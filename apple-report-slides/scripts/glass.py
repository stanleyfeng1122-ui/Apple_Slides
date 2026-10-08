"""Verified glass pipeline: photo mask, rim note, tint calibration, 2x face bake.

Port of the round-5 prototype. Optical attribution is in assets/THIRD_PARTY.md.
No private paths or case content; one browser page is shared with the converter.
"""
import sys
sys.dont_write_bytecode = True

import base64
from io import BytesIO
import math
from pathlib import Path
import re
from urllib.parse import unquote, urlparse

import numpy as np
from PIL import Image, ImageDraw
from pptx.oxml.xmlchemy import OxmlElement

WARNING = ('Glass is a baked picture: rebuild the deck after moving it, editing '
           'what is under it, or lengthening its text.')
PICTURE_NAME = 'glass picture: rebuild if moved or content beneath changes'
PAD = 1
CAP = 0.25
# Mean |L - blur(L)| (0-255) in a box's rim band. Foliage, mesh and gravel sit
# above it; the approved prototype faces (C1-C4) sit at or below 10.
FINE_TEXTURE = 12
FINE_FRINGE = 0.6
_fringes = {}  # one decision per box, shared by the 1x measurement and 2x capture pages


def local_image(url):
    parsed = urlparse(url)
    if parsed.scheme == 'data' and ';base64,' in url:
        return Image.open(BytesIO(base64.b64decode(url.split(',', 1)[1])))
    if parsed.scheme == 'file' and parsed.netloc in ('', 'localhost'):
        return Image.open(Path(unquote(parsed.path)).resolve(strict=True))
    raise ValueError('Bento photos must be local files or embedded base64 images')


def data_url(image, kind='PNG'):
    buf = BytesIO()
    image.save(buf, kind, **({'quality': 93, 'subsampling': 0} if kind == 'JPEG' else {}))
    mime = 'jpeg' if kind == 'JPEG' else 'png'
    return f'data:image/{mime};base64,' + base64.b64encode(buf.getvalue()).decode()


def bento_mask(box, tiles):
    """The prototype's tile union + interior + round gutter tips, from page geometry."""
    ox, oy, w, h = box
    radius, gutter, ss = 40, 24, 4
    mask = Image.new('L', (round(w * ss), round(h * ss)), 0)
    draw = ImageDraw.Draw(mask)

    def q(x0, y0, x1, y1):
        return ((x0 - ox) * ss, (y0 - oy) * ss,
                (x1 - ox) * ss - 1, (y1 - oy) * ss - 1)

    for x, y, tw, th in tiles:
        draw.rounded_rectangle(q(x, y, x + tw, y + th), radius=radius * ss, fill=255)
    k = gutter / 2
    draw.rectangle(q(ox + radius, oy + radius + k,
                     ox + w - radius - k, oy + h - radius - k), fill=255)
    stubs = set()
    # Find neighbouring tile pairs on each outside edge. Interior gaps remain photo.
    for x, y, tw, th in tiles:
        for bx, by, bw, bh in tiles:
            if abs(bx - x - tw - gutter) < .1:
                if abs(y - oy) < .1 and abs(by - oy) < .1:
                    stubs.add((x + tw + k, oy + radius, 0, -1))
                if abs(y + th - oy - h) < .1 and abs(by + bh - oy - h) < .1:
                    stubs.add((x + tw + k, oy + h - radius, 0, 1))
            if abs(by - y - th - gutter) < .1:
                if abs(x - ox) < .1 and abs(bx - ox) < .1:
                    stubs.add((ox + radius, y + th + k, -1, 0))
                if abs(x + tw - ox - w) < .1 and abs(bx + bw - ox - w) < .1:
                    stubs.add((ox + w - radius, y + th + k, 1, 0))
    for cx, cy, dx, dy in stubs:
        ex, ey = cx - dx * k, cy - dy * k
        draw.ellipse(q(ex - k, ey - k, ex + k, ey + k), fill=255)
        draw.rectangle(q(min(ex, ex - dx * radius) - k * (dx == 0),
                         min(ey, ey - dy * radius) - k * (dy == 0),
                         max(ex, ex - dx * radius) + k * (dx == 0),
                         max(ey, ey - dy * radius) + k * (dy == 0)), fill=255)
    return mask.resize((round(w), round(h)), Image.Resampling.LANCZOS)


def auto_fringe(page, report=True):
    """Measure rim-band fine texture under each box; .6 fringe where it would sparkle."""
    from PIL import ImageFilter
    ids = page.locator('.page').evaluate_all('ps=>ps.map(p=>p.id)')
    yy, xx = np.mgrid[0:1080, 0:1920] + .5
    for pid in ids:
        current = page.locator(f'[id="{pid}"]')
        boxes = current.evaluate("""p=>[...p.querySelectorAll('[data-role="glass"]')].map(g=>{
          const s=getComputedStyle(g);return {name:g.dataset.name||'glass',fringe:g.dataset.fringe||null,
          box:['left','top','width','height','borderTopLeftRadius'].map(k=>parseFloat(s[k])||0)}})""")
        if not boxes:
            continue
        key = (page.url, pid, str(boxes))
        if key in _fringes:
            current.evaluate("""(p,f)=>p.querySelectorAll('[data-role="glass"]').forEach((g,i)=>{if(f[i]) g.dataset.fringe=f[i]})""", _fringes[key])
            continue
        page.evaluate('(id)=>showPage(id)', pid)
        current.evaluate("p=>p.querySelectorAll('[data-role=glass]').forEach(g=>g.style.visibility='hidden')")
        try:
            grey = screenshot(page).convert('L')
        finally:
            current.evaluate("p=>p.querySelectorAll('[data-role=glass]').forEach(g=>g.style.removeProperty('visibility'))")
        detail = np.abs(np.asarray(grey, float) - np.asarray(grey.filter(ImageFilter.GaussianBlur(2)), float))
        chosen = []
        for item in boxes:
            x, y, w, h, r = item['box']
            r = min(r, w/2, h/2)
            bezel = min(max(.3 * min(w, h), 16), 56)
            qx, qy = np.abs(xx-x-w/2)-(w/2-r), np.abs(yy-y-h/2)-(h/2-r)
            distance = np.hypot(np.maximum(qx, 0), np.maximum(qy, 0)) + np.minimum(np.maximum(qx, qy), 0) - r
            energy = float(detail[(distance >= -bezel) & (distance <= 0)].mean())
            if item['fringe'] is not None:
                fringe, why = float(item['fringe']), 'authored'
            else:
                fringe = FINE_FRINGE if energy > FINE_TEXTURE else 1.0
                why = 'fine texture' if fringe == FINE_FRINGE else 'default'
            chosen.append(None if why == 'authored' else str(fringe))
            if report:
                print(f"{pid} · {item['name']} · rim detail {energy:.1f} · fringe {fringe:g} ({why})", flush=True)
        _fringes[key] = chosen
        current.evaluate("""(p,f)=>p.querySelectorAll('[data-role="glass"]').forEach((g,i)=>{if(f[i]) g.dataset.fringe=f[i]})""", chosen)


def prepare(page, report=True):
    """Prepare bento photo masks before the unchanged lens builds its backdrop clones."""
    page._glass_photos = {}
    photos = page.locator('.page > .backdrop [data-bento-bleed]')
    for index in range(photos.count()):
        photo = photos.nth(index)
        data = photo.evaluate('''e => {
          const box = e => {const s=getComputedStyle(e);return [s.left,s.top,s.width,s.height].map(parseFloat)};
          return {id:e.closest('.page').id,src:new URL(e.dataset.bentoBleed,document.baseURI).href, box:box(e),
            bleed:e.dataset.lensBox.split(',').map(Number),
            tiles:Array.from(e.closest('.page').querySelectorAll('[data-role="glass"]')).map(box)};
        }''')
        if not data['tiles']:
            raise ValueError('A bento photo needs glass tile geometry')
        x, y, w, h = data['box']
        bx, by, bw, bh = data['bleed']
        data['real_bleed'] = list(data['bleed'])
        with local_image(data['src']) as source:
            if source.width < bw or source.height < bh:
                raise ValueError('Bento photo is too small: use a sharp native crop without upscaling')
            bleed = source.convert('RGB').resize((round(bw), round(bh)), Image.Resampling.LANCZOS)
        data['whole_image'] = bleed.copy()
        # Padding belongs to the lens only. Keep the real crop bounds separately
        # so a placement search can never reveal reflected pixels on the page.
        left, top = max(70, math.ceil(70-(x-bx))), max(70, math.ceil(70-(y-by)))
        right, bottom = max(70, math.ceil(70-(bx+bw-x-w))), max(70, math.ceil(70-(by+bh-y-h)))
        if any((left, top, right, bottom)):
            bleed = Image.fromarray(np.pad(np.asarray(bleed), ((top,bottom),(left,right),(0,0)), mode='reflect'))
            bx, by, bw, bh = bx-left, by-top, bw+left+right, bh+top+bottom
            data['bleed'] = [bx, by, bw, bh]
            if report:
                print(f"{data['id']} · lens source mirror-padded {left},{top},{right},{bottom} px", flush=True)
        data['image'] = bleed
        page._glass_photos[data['id']] = data
        clipped = bleed.crop((round(x-bx), round(y-by), round(x-bx+w), round(y-by+h))).convert('RGBA')
        clipped.putalpha(bento_mask(data['box'], data['tiles']))
        photo.evaluate('''(e, args) => {
          e.src=args[0];e.dataset.lensSrc=args[1];e.dataset.lensBox=args[2];e.style.objectFit='fill';
        }''', [data_url(clipped), data_url(bleed), ','.join(map(str,data['bleed']))])
    if page.locator('[data-role="glass"]').count():
        if not page.evaluate('window.glassReady'):
            page.evaluate('async()=>{await Promise.all(Array.from(document.images).map(i=>i.decode()))}')
            auto_fringe(page, report)
        page.evaluate('''async () => {
          await Promise.all(Array.from(document.images).map(i=>i.decode()));
          if (!window.buildGlass) throw new Error('Glass runtime missing: copy glass.js beside the deck');
          if (!window.glassReady) {window.buildGlass();window.glassReady=true;}
          await Promise.all(Array.from(document.images).map(i=>i.decode()));
        }''')


def luminance(rgb):
    v = np.asarray(rgb, dtype=float) / 255
    linear = np.where(v <= .04045, v / 12.92, ((v + .055) / 1.055) ** 2.4)
    return .2126 * linear[..., 0] + .7152 * linear[..., 1] + .0722 * linear[..., 2]


def screenshot(page):
    # All measurements use 1x coordinates, even when final baking is 2x.
    image = Image.open(BytesIO(page.screenshot())).convert('RGB')
    if image.size != (1920, 1080):
        image = image.resize((1920, 1080), Image.Resampling.LANCZOS)
    return image


MEASURE = '''g => {
  const texts=[];
  const inks=JSON.parse(g.dataset.measureColors||'[]');
  g.querySelectorAll(':scope > [data-role="text"]').forEach(t=>{
    const s=getComputedStyle(t), walker=document.createTreeWalker(t,NodeFilter.SHOW_TEXT);
    let node;
    while(node=walker.nextNode()) {
      if(!node.textContent.trim()) continue;
      const r=document.createRange();r.selectNodeContents(node);
      const ts=getComputedStyle(node.parentElement);
      for(const b of r.getClientRects()) if(b.width && b.height)
        texts.push({text:node.textContent.trim(), color:inks.find(t=>t.text===node.textContent.trim())?.color||ts.color, size:parseFloat(ts.fontSize),
          box:[b.x,b.y,b.width,b.height]});
    }
  });
  const s=getComputedStyle(g), r=g.getBoundingClientRect();
  return {name:g.dataset.name||'glass',box:[r.x,r.y,r.width,r.height],
    radius:parseFloat(s.borderTopLeftRadius),optics:JSON.parse(g.dataset.optics),
    base:parseFloat(g.dataset.tintBase ?? g.dataset.tint ?? (g.classList.contains('dark') ? '.14' : '.12')),
    fixedBase:g.dataset.tintBase!==undefined || g.dataset.tint!==undefined,
    tint:parseFloat(g.dataset.tint ?? s.getPropertyValue('--a')),texts};
}'''


def rim_note(page, current, glasses, data):
    current.evaluate("p=>p.querySelectorAll('[data-role=glass]').forEach(g=>g.style.visibility='hidden')")
    try:
        lum = luminance(np.asarray(screenshot(page)))
    finally:
        current.evaluate("p=>p.querySelectorAll('[data-role=glass]').forEach(g=>g.style.removeProperty('visibility'))")
    light = np.where(lum > .008856, 116 * np.cbrt(lum) - 16, 903.3 * lum)
    yy, xx = np.mgrid[0:1080, 0:1920] + .5
    for item in data:
        x, y, w, h = item['box']
        r = min(item['radius'], w/2, h/2)
        qx, qy = np.abs(xx-x-w/2)-(w/2-r), np.abs(yy-y-h/2)-(h/2-r)
        distance = np.hypot(np.maximum(qx, 0), np.maximum(qy, 0)) + np.minimum(np.maximum(qx, qy), 0) - r
        pixels = light[(distance >= -6) & (distance <= item['optics']['pull'])]
        delta = float(np.percentile(pixels, 95) - np.percentile(pixels, 5))
        print(f"{current.get_attribute('id')} · {item['name']} · rim ΔL* {delta:.1f} (information only)", flush=True)


def contrast(image, text):
    x, y, w, h = text['box']
    pixels = np.asarray(image.crop((max(0, int(x)-4), max(0, int(y)-2),
                                   min(1920, int(x+w)+4), min(1080, int(y+h)+2))))
    lt = float(luminance([float(n) for n in re.findall(r'[\d.]+', text['color'])[:3]]))
    worst = float(np.percentile(luminance(pixels), 90 if lt > .5 else 10))
    return (max(lt, worst)+.05)/(min(lt, worst)+.05)


def set_photo_offset(page, current, offset):
    """Move the shared source and the page crop together, retaining lens bleed."""
    info = page._glass_photos.get(current.get_attribute('id'))
    if not info:
        return
    x, y, w, h = info['box']
    bx, by, bw, bh = info['bleed']
    bx, by = bx+offset[0], by+offset[1]
    clipped = info['image'].crop((round(x-bx),round(y-by),round(x-bx+w),round(y-by+h))).convert('RGBA')
    clipped.putalpha(bento_mask(info['box'], info['tiles']))
    current.evaluate('''(p,a) => {
      const photo=p.querySelector(':scope > .backdrop [data-bento-bleed]');
      photo.src=a.src;photo.dataset.lensBox=a.box.join(',');
      p.querySelectorAll('.optic [data-bento-bleed]').forEach(e=>{
        const [x,y,w,h]=a.box;Object.assign(e.style,{left:x+'px',top:y+'px',width:w+'px',height:h+'px'});
      });
    }''', dict(src=data_url(clipped),box=[bx,by,bw,bh]))
    page.evaluate('async()=>{await Promise.all(Array.from(document.images).map(i=>i.decode()))}')


PLACE = '''(g,position) => {
  const texts=[...g.querySelectorAll(':scope > [data-role="text"]')];
  if(!g.dataset.textOrigins) g.dataset.textOrigins=JSON.stringify(texts.map(t=>t.getAttribute('style')));
  if(position==='original') {
    JSON.parse(g.dataset.textOrigins).forEach((s,i)=>texts[i].setAttribute('style',s));return;
  }
  const s=getComputedStyle(g),w=parseFloat(s.width),h=parseFloat(s.height);
  const pad=Math.min(parseFloat(s.borderTopLeftRadius)||40,44), width=w-pad*2;
  const sorted=texts.slice().sort((a,b)=>parseFloat(getComputedStyle(a).top)-parseFloat(getComputedStyle(b).top));
  for(const t of sorted) Object.assign(t.style,{left:pad+'px',width:width+'px',height:'auto',
    textAlign:position.endsWith('right')?'right':'left'});
  const heights=sorted.map(t=>Math.max(t.getBoundingClientRect().height,parseFloat(getComputedStyle(t).lineHeight)));
  const total=heights.reduce((a,b)=>a+b,0)+8*(heights.length-1);
  let top=position.startsWith('bottom')?h-pad-total:position.startsWith('centre')?(h-total)/2:pad;
  sorted.forEach((t,i)=>{t.style.top=top+'px';top+=heights[i]+8});
}'''


def place(glasses, positions):
    for i, position in enumerate(positions):
        glasses.nth(i).evaluate(PLACE, position)


def set_tones(page, current, tones):
    """Use the existing clear/light material on bright photo zones, smoke on dark."""
    current.evaluate('''(p,tones)=>p.querySelectorAll('[data-role="glass"]').forEach((g,i)=>{
      const tone=tones[i];g.classList.remove('dark','light');g.classList.add(tone);
      const optics=JSON.parse(g.dataset.optics);optics.tone=tone;g.dataset.optics=JSON.stringify(optics);
      const backdrop=g.querySelector('.optic > .backdrop');
      backdrop.querySelectorAll('[data-lens-shadow]').forEach(e=>e.remove());
      [...backdrop.children].filter(e=>!e.dataset.role && getComputedStyle(e).boxShadow!=='none').forEach(e=>e.remove());
      if(tone==='light') {
        const cs=getComputedStyle(g),shadow=document.createElement('div');shadow.dataset.lensShadow='true';
        Object.assign(shadow.style,{position:'absolute',left:cs.left,top:cs.top,width:cs.width,height:cs.height,
          borderRadius:cs.borderTopLeftRadius,boxShadow:cs.boxShadow});backdrop.append(shadow);
      }
      delete g.dataset.measureColors;
    })''', tones)
    measuring=page.evaluate("document.documentElement.classList.contains('glass-measure')")
    page.evaluate("document.documentElement.classList.remove('glass-measure')")
    glasses=current.locator('[data-role="glass"]')
    inks=[[dict(text=t['text'],color=t['color']) for t in glasses.nth(i).evaluate(MEASURE)['texts']]
          for i in range(glasses.count())]
    current.evaluate('(p,inks)=>p.querySelectorAll("[data-role=glass]").forEach((g,i)=>g.dataset.measureColors=JSON.stringify(inks[i]))',inks)
    if measuring:
        page.evaluate("document.documentElement.classList.add('glass-measure')")


def measure_tints(page, glasses):
    """Find the smallest passing hundredth, from the prescribed floor to .25."""
    data = [glasses.nth(i).evaluate(MEASURE) for i in range(glasses.count())]
    low = [max(-1, math.ceil(d['base']*100-1e-8)-1) for d in data]
    high = [26]*len(data)
    probe = [value+1 for value in low]
    first = True
    last = [0.0]*len(data)
    while any(h-l>1 for l,h in zip(low, high)):
        for i, value in enumerate(probe):
            glasses.nth(i).evaluate("(g,a)=>g.style.setProperty('--a',a)", str(min(value,25)/100))
        frame = screenshot(page)
        for i, item in enumerate(data):
            x,y,w,h=item['box']
            fits=all(t['box'][0]>=x+4 and t['box'][1]>=y+4 and
                     t['box'][0]+t['box'][2]<=x+w-4 and t['box'][1]+t['box'][3]<=y+h-4 for t in item['texts'])
            ratios = [contrast(frame,t)/(3.5 if t['size']>=40 else 5.0) for t in item['texts']]
            last[i] = min(ratios, default=1)
            if not fits:
                last[i]=0
            if high[i]-low[i]<=1:
                continue
            if last[i]>=1:
                high[i]=probe[i]
            else:
                low[i]=probe[i]
        # Try the cap immediately after the floor; don't spend several renders
        # bisecting a box for which even .25 cannot pass.
        probe = [25 if first and h==26 and h-l>1 else
                 min(25,(l+h)//2) if h-l>1 else min(h,25) for l,h in zip(low,high)]
        first = False
    return [v/100 for v in high]


def plain_photo_fallback(page, current):
    """One whole uncovered photo plus native text tiles; no photographic gutters."""
    info = page._glass_photos.get(current.get_attribute('id'))
    whole = data_url(info['whole_image']) if info else None
    current.evaluate('''(p,whole) => {
      const photo=p.querySelector(':scope > .backdrop img');
      if(!photo) throw new Error('Glass fallback requires a photo; move this callout off the chart');
      if(whole) photo.src=whole;
      delete photo.dataset.bentoBleed;delete photo.dataset.lensSrc;delete photo.dataset.lensBox;
      Object.assign(photo.style,{left:'90px',top:'250px',width:'858px',height:'680px',objectFit:'contain',borderRadius:'40px'});
      const boxes=[...p.querySelectorAll('[data-role="glass"]')], columns=boxes.length>3?2:1;
      const rows=Math.ceil(boxes.length/columns),width=(858-24*(columns-1))/columns,height=(680-24*(rows-1))/rows;
      boxes.forEach((g,i)=>{
        g.querySelectorAll('.lens,.tint,.rim').forEach(e=>e.remove());
        [...p.querySelectorAll(':scope > [data-glass-for]')].filter(t=>t.dataset.glassFor===g.dataset.name).forEach(t=>g.append(t));
        const texts=[...g.querySelectorAll(':scope > [data-role="text"]')];
        g.className='card';g.removeAttribute('data-role');
        Object.assign(g.style,{left:(972+(i%columns)*(width+24))+'px',top:(250+Math.floor(i/columns)*(height+24))+'px',
          width:width+'px',height:height+'px',boxShadow:'none'});
        const tile=document.createElement('div');tile.dataset.role='shape';tile.className='bento-tile';tile.style.inset='0';g.prepend(tile);
        let y=32;
        texts.forEach(t=>{
          const old=getComputedStyle(t),value=parseFloat(old.fontSize)>=40;
          Object.assign(t.style,{left:'32px',top:y+'px',width:(width-64)+'px',height:'auto',textAlign:'left',color:'var(--ink)',textShadow:'none'});
          t.style.setProperty('font-size',(value?40:20)+'px','important');
          t.style.setProperty('line-height',(value?48:28)+'px','important');
          y+=t.getBoundingClientRect().height+8;
        });
      });
      const captions=[...p.children].filter(e=>e.dataset.role==='text' &&
        !e.matches('.headline,.eyebrow,.subtitle,.footer,.footer-source,.footer-mark,.footer-number'));
      captions.forEach((t,i)=>Object.assign(t.style,{left:(90+i*1740/captions.length)+'px',top:'944px',
        width:(1740/captions.length-16)+'px',height:'auto',color:'var(--ink)',textShadow:'none'}));
      p.dataset.layout='bento-plain';p.dataset.variant='with-photo-tile';p.dataset.glassFallback='cap';
    }''', whole)
    print(f"{current.get_attribute('id')} · FALLBACK: whole uncovered photo tile + bento-plain", flush=True)


def apply_selection(page, current, selection):
    if not selection:
        return
    if selection.get('fallback'):
        plain_photo_fallback(page,current)
        return
    set_photo_offset(page,current,selection['offset'])
    if selection.get('tones'):
        set_tones(page,current,selection['tones'])
    glasses=current.locator('[data-role="glass"]')
    place(glasses,selection['positions'])
    for i,tint in enumerate(selection['tints']):
        glasses.nth(i).evaluate("(g,a)=>g.style.setProperty('--a',a)",str(tint))


def calibrate(page, current, *, report_rims=True):
    """Search text zones and available 40 px photo offsets before a whole-photo fallback."""
    glasses=current.locator('[data-role="glass"]')
    data=[glasses.nth(i).evaluate(MEASURE) for i in range(glasses.count())]
    if not data:
        return None
    if any(not math.isfinite(d['base']) or not 0<=d['base']<=CAP for d in data):
        raise ValueError('Glass tint floor must be between 0 and .25')
    place(glasses,['original']*len(data))
    page.add_style_tag(content='html.glass-measure .glass > [data-role="text"], html.glass-measure .glass > [data-role="text"] * {color:transparent!important}')
    page.evaluate("document.documentElement.classList.add('glass-measure')")
    # Store real foreground colours before the measurement CSS hides the glyphs.
    try:
        # Color-transparent measurement keeps its halo but needs the actual ink
        # recorded before hiding. Save it on the element for MEASURE below.
        current.evaluate('(p,colors)=>p.querySelectorAll("[data-role=glass]").forEach((g,i)=>g.dataset.measureColors=JSON.stringify(colors[i]))',
                         [[dict(text=t['text'],color=t['color']) for t in d['texts']] for d in data])
        tints=measure_tints(page,glasses)
        original_tones=[d['optics']['tone'] for d in data]
        best=dict(tints=tints,positions=['original']*len(data),offset=[0,0],tones=original_tones,fallback=False)
        def score(values):
            return (sum(v>CAP for v in values),max(values),sum(values))
        if any(t>CAP for t in tints):
            pid=current.get_attribute('id')
            print(f'{pid} · searching five text zones and available photo bleed',flush=True)
            info=page._glass_photos.get(pid)
            offsets=[(0,0)]
            if info:
                x,y,w,h=info['box'];bx,by,bw,bh=info['real_bleed']
                # Search the real crop's full bleed in 40 px steps, plus its
                # original alignment. Mirror pixels remain lens-only.
                xs=sorted({0,*range(math.ceil(x+w-bx-bw),math.floor(x-bx)+1,40)})
                ys=sorted({0,*range(math.ceil(y+h-by-bh),math.floor(y-by)+1,40)})
                offsets=sorted(((dx,dy) for dx in xs for dy in ys),key=lambda o:abs(o[0])+abs(o[1]))
            for offset in offsets:
                set_photo_offset(page,current,offset)
                local_tints=[.26]*len(data);positions=['original']*len(data);tones=list(original_tones)
                variants=[original_tones]
                if info:
                    variants.append(['light' if tone=='dark' else 'dark' for tone in original_tones])
                for variant in variants:
                    set_tones(page,current,variant)
                    for position in ['top-left','bottom-left','bottom-right','top-right','centre-left']:
                        place(glasses,[position]*len(data))
                        candidates=measure_tints(page,glasses)
                        for i,tint in enumerate(candidates):
                            if tint<local_tints[i]:
                                local_tints[i]=tint;positions[i]=position;tones[i]=variant[i]
                if score(local_tints)<score(best['tints']):
                    best=dict(tints=local_tints,positions=positions,offset=list(offset),tones=tones,fallback=False)
                if all(t<=math.ceil((d['base'] if d['fixedBase'] else min(d['base'],.12))*100-1e-8)/100
                       for t,d in zip(best['tints'],data)):
                    break  # The prescribed floor is the global minimum.
        if any(t>CAP for t in best['tints']):
            print(f"{current.get_attribute('id')} · CAP HIT after all placements/offsets",flush=True)
            best['fallback']=True
    finally:
        page.evaluate("document.documentElement.classList.remove('glass-measure')")
    apply_selection(page,current,{**best,'fallback':False,'tints':[min(t,CAP) for t in best['tints']]})
    final=[glasses.nth(i).evaluate(MEASURE) for i in range(glasses.count())]
    if report_rims:
        rim_note(page,current,glasses,final)
    for item,tint,position,tone in zip(final,best['tints'],best['positions'],best['tones']):
        tint_text=f'{tint:.2f}' if tint<=CAP else '>.25 (no passing placement)'
        print(f"{current.get_attribute('id')} · {item['name']} · tint {tint_text} · "
              f"placement {position} · tone {tone} · photo offset {best['offset'][0]},{best['offset'][1]} px",flush=True)
    if best['fallback']:
        plain_photo_fallback(page,current)
    return best


def bake(page, current, data):
    """Replace only glass items with individually captured JPEGs; text items are untouched."""
    items = [item for item in data['elements'] if item['role'] == 'glass']
    if not items:
        return
    page.add_style_tag(content='''html.glass-capture .glass {box-shadow:none!important}
      html.glass-capture .glass > [data-role="text"],
      html.glass-capture .glass > [data-role="text"] * {color:transparent!important}''')
    glasses = current.locator('[data-role="glass"]')
    shadows = [glasses.nth(i).evaluate('g=>getComputedStyle(g).boxShadow') for i in range(len(items))]
    page.evaluate("document.documentElement.classList.add('glass-capture')")
    try:
        for item, shadow in zip(items, shadows):
            b = item['box']
            box = dict(x=b['x']-PAD, y=b['y']-PAD, w=b['w']+2*PAD, h=b['h']+2*PAD)
            face = Image.open(BytesIO(page.screenshot(clip=dict(x=box['x'], y=box['y'],
                                    width=box['w'], height=box['h'])))).convert('RGB')
            item.update(role='image', fit='fill', src=data_url(face, 'JPEG'), box=box,
                        name=f"{item['name']} ({PICTURE_NAME})",
                        glass=dict(radius=min(item['style']['radius']+PAD, box['w']/2, box['h']/2), shadow=shadow))
    finally:
        page.evaluate("document.documentElement.classList.remove('glass-capture')")
    notes = data['notes'].rstrip()
    old_warning = ('Glass is a baked picture: rebuild the deck after moving it '
                   'or editing what is under it.')
    if notes.endswith(old_warning):
        notes = notes[:-len(old_warning)].rstrip()
    if not notes.endswith(WARNING):
        notes += '\n' + WARNING
    data['notes'] = notes


def finish_picture(shape, item):
    """Native roundRect picture crop and editable outer shadow, as in the prototype."""
    info, b = item['glass'], item['box']
    props = shape._element.spPr
    geom = props.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
    geom.set('prst', 'roundRect')
    av = geom.find('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst')
    for old in list(av):
        av.remove(old)
    adjustment = OxmlElement('a:gd')
    adjustment.set('name', 'adj')
    adjustment.set('fmla', f"val {round(min(50000, info['radius']/min(b['w'], b['h'])*100000))}")
    av.append(adjustment)
    match = re.match(r'rgba?\(([^)]*)\)\s+(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px', info['shadow'])
    if match:
        rgba = [float(v) for v in match.group(1).split(',')]
        dx, dy, blur = float(match.group(2)), float(match.group(3)), float(match.group(4))
        effects = OxmlElement('a:effectLst')
        shadow = OxmlElement('a:outerShdw')
        for key, value in dict(blurRad=round(blur*6350), dist=round(math.hypot(dx,dy)*6350),
                               dir=round(math.degrees(math.atan2(dy,dx))*60000) % 21600000, algn='t', rotWithShape='0').items():
            shadow.set(key, str(value))
        colour = OxmlElement('a:srgbClr')
        colour.set('val', ''.join(f'{round(v):02X}' for v in rgba[:3]))
        alpha = OxmlElement('a:alpha')
        alpha.set('val', str(round((rgba[3] if len(rgba)>3 else 1)*100000)))
        colour.append(alpha)
        shadow.append(colour)
        effects.append(shadow)
        props.append(effects)
