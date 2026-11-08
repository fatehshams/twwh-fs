# Usage: python3 make_cover.py MedulaOne-Regular.ttf out.jpg 330 220 255 0.78 1  (font size, top y, line gap, width squeeze, stroke)
import cv2, numpy as np, sys
from PIL import Image, ImageDraw, ImageFont, ImageFilter
D='/tmp/claude-0/-home-user-twwh-fs/d31a081f-9d13-5635-acab-4f78162195b0/scratchpad/'
img=cv2.imread(D+'ref/apple.jpg')
hsv=cv2.cvtColor(img,cv2.COLOR_BGR2HSV)
mask=np.zeros(img.shape[:2],np.uint8)
light=((hsv[...,1]<90)&(hsv[...,2]>150)).astype(np.uint8)*255
mask[150:870,90:690]=light[150:870,90:690]          # title
mask[1170:1310,100:690]=light[1170:1310,100:690]    # ENGLISH
mask=cv2.dilate(mask,np.ones((9,9),np.uint8),2)
clean=cv2.inpaint(img,mask,12,cv2.INPAINT_TELEA)
# smooth the inpainted zones to kill smudges
blur=cv2.GaussianBlur(clean,(0,0),40)
m3=cv2.GaussianBlur(cv2.dilate(mask,np.ones((61,61),np.uint8)),(0,0),20)[...,None]/255.
clean=(clean*(1-m3)+blur*m3).astype(np.uint8)
im=Image.fromarray(cv2.cvtColor(clean,cv2.COLOR_BGR2RGB)).convert('RGBA')
SQ=float(sys.argv[6]); SW=int(sys.argv[7]); font=sys.argv[1]; out=sys.argv[2]
cream=(236,226,218,255)
def text(lines,fnt,x,y0,gap,fill,shadow=True):
    lay=Image.new('RGBA',im.size,(0,0,0,0)); d=ImageDraw.Draw(lay)
    y=y0
    for l in lines:
        w=int(d.textlength(l,font=fnt))+40; h=fnt.size*2
        t=Image.new('RGBA',(w,h),(0,0,0,0)); ImageDraw.Draw(t).text((0,0),l,font=fnt,fill=fill,stroke_width=SW,stroke_fill=fill)
        t=t.resize((int(w*SQ),h),Image.LANCZOS); lay.alpha_composite(t,(x,y)); y+=gap
    if shadow:
        sh=Image.new('RGBA',im.size,(0,0,0,0)); sh.putalpha(lay.getchannel('A').filter(ImageFilter.GaussianBlur(6)).point(lambda a:a*0.45))
        im.alpha_composite(sh,(4,6))
    im.alpha_composite(lay)
f=ImageFont.truetype(D+'fonts/'+font,int(sys.argv[3]))
text(['Fredags-','predikan'],f,105,int(sys.argv[4]),int(sys.argv[5]),cream)
bold=ImageFont.truetype(D+'fonts/Oswald.ttf',150); bold.set_variation_by_axes([600])
lay=Image.new('RGBA',im.size,(0,0,0,0)); d=ImageDraw.Draw(lay)
s='SVENSKA'; x=112
for ch in s:  # letter-spaced like ENGLISH
    d.text((x,1150),ch,font=bold,fill=(255,255,255,255)); x+=d.textlength(ch,font=bold)+14
im.alpha_composite(lay)
im=im.convert('RGB').resize((3000,3000),Image.LANCZOS)
im.save(D+'out/'+out,quality=95)
