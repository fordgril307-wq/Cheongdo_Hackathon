import json, math, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
fp="/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"; fb="/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
fm.fontManager.addfont(fp); fm.fontManager.addfont(fb)
plt.rcParams['font.family']=fm.FontProperties(fname=fp).get_name()
d=json.load(open('muni.json')); F={f['properties']['code']:f for f in d['features']}
def polys(f):
    g=f['geometry']; cs=g['coordinates']
    return [p[0] for p in cs] if g['type']=='MultiPolygon' else [cs[0]]
def draw(ax,code,fc,ec,lw=1,z=1):
    for ring in polys(F[code]):
        ax.fill([c[0] for c in ring],[c[1] for c in ring],fc=fc,ec=ec,lw=lw,zorder=z)
NAVY="#0B1F4D"; OR="#C2410C"; G="#8C95A6"
def base(xl,yl,size):
    fig,ax=plt.subplots(figsize=size,dpi=200)
    for c in ['37100','37070','37020','38080','22310']: draw(ax,c,"#F4F5F7","#FFFFFF",0.8,0)
    draw(ax,'37360',"#FBF7F2","#B9A58F",1.4,1)
    ax.set_xlim(*xl); ax.set_ylim(*yl); ax.set_aspect(1/0.81); ax.axis('off'); return fig,ax
def dist(a,b):
    k=math.hypot((b[0]-a[0])*90.6,(b[1]-a[1])*111)*1.35; return k, round(k/40*60)
S=(128.7233,35.6858)
pts={"청도역·상상마루":(128.7466,35.6405),"와인터널":(128.7205,35.7151),"다로리 마을호텔\n·마을영화관(2027)":(128.7213,35.7068),
     "신화랑풍류마을":(128.9223,35.7181),"운문사":(128.9566,35.6740)}
# --- 입지 map (radial)
fig,ax=base((128.60,129.0),(35.60,35.76),(8,4.2))
lab={"청도역·상상마루":(0.006,-0.012),"와인터널":(0.006,0.004),"다로리 마을호텔\n·마을영화관(2027)":(-0.105,-0.002),"신화랑풍류마을":(-0.02,0.008),"운문사":(0.006,-0.012)}
for k,p in pts.items():
    km,mn=dist(S,p)
    ax.plot([S[0],p[0]],[S[1],p[1]],color="#E8A87C",lw=1.4,ls=(0,(3,2)),zorder=2)
    ax.plot([p[0]],[p[1]],'o',ms=9,color=NAVY,mec='white',mew=1.2,zorder=5)
    ax.text(p[0]+lab[k][0],p[1]+lab[k][1],k,fontsize=8.5,color=NAVY,weight='bold',zorder=6)
    mx,my=(S[0]+p[0])/2,(S[1]+p[1])/2
    if k=="와인터널": mx+=0.02
    if k.startswith("다로리"): mx-=0.035; my-=0.003
    ax.text(mx,my,f"{km:.0f}km·{mn}분",fontsize=7.5,color="#7A2E0B",ha='center',va='center',zorder=7,bbox=dict(boxstyle="round,pad=0.2",fc="#FFF4EC",ec="#F2C9A8",lw=0.6))
ax.plot([S[0]],[S[1]],'*',ms=20,color=OR,mec='white',mew=1,zorder=8)
ax.text(S[0]-0.006,S[1]-0.016,"청도 소싸움경기장\n(돔, 약 1만1천 석)",fontsize=10,color=OR,weight='bold',ha='right',zorder=8)
ax.text(128.605,35.605,"거리: 직선×1.35 · 시간: 시속 40km 추정 · 동대구역까지 약 31km·47분(추정)",fontsize=7,color=G)
plt.savefig('locmap.png',bbox_inches='tight',pad_inches=0.05,facecolor='white')
# --- course map
fig,ax=base((128.6,129.06),(35.6,35.76),(8,3.9))
P={"청도역·상상마루":(128.7466,35.6405),"와인터널":(128.7205,35.7151),"소싸움경기장":S,"신화랑풍류마을":(128.9223,35.7181),"운문사":(128.9566,35.6740)}
route=list(P.keys()); LOFF=[(0.035,-0.015),(-0.04,-0.002),(0.0,-0.012),(0.012,0.002)]
for i,(a,b) in enumerate(zip(route,route[1:])):
    x1,y1=P[a]; x2,y2=P[b]; km,mn=dist(P[a],P[b])
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),arrowprops=dict(arrowstyle="-|>",color=OR,lw=1.8,shrinkA=7,shrinkB=7,connectionstyle="arc3,rad=0.15"),zorder=3)
    ax.text((x1+x2)/2+LOFF[i][0],(y1+y2)/2+LOFF[i][1],f"약 {km:.0f}km·{mn}분",fontsize=7.5,color="#7A2E0B",ha='center',va='center',zorder=8,bbox=dict(boxstyle="round,pad=0.2",fc="#FFF4EC",ec="#F2C9A8",lw=0.6))
off={"청도역·상상마루":(0.008,-0.012),"와인터널":(0.008,0.004),"소싸움경기장":(-0.07,-0.004),"신화랑풍류마을":(-0.02,0.009),"운문사":(0.008,-0.012)}
for i,(k,(x,y)) in enumerate(P.items()):
    ax.plot([x],[y],'o',ms=13,color=OR,mec='white',mew=1.2,zorder=5)
    ax.text(x,y,str(i+1),fontsize=7,color='white',weight='bold',ha='center',va='center',zorder=7)
    ax.text(x+off[k][0],y+off[k][1],k,fontsize=9,weight='bold',color=OR,zorder=6)
dx,dy=(128.7213,35.7068); ax.plot([dx],[dy],'s',ms=8,color=NAVY,mec='white',zorder=6); ax.text(dx-0.075,dy+0.006,"다로리(2027)",fontsize=8,color=NAVY,zorder=6)
for n,(x,y) in {"대천터미널":(128.9143,35.7164),"동곡터미널":(128.8907,35.6934)}.items():
    ax.plot([x],[y],'D',ms=6,color=G,zorder=4); ax.text(x-0.03,y-0.014,n,fontsize=7.5,color=G)
plt.savefig('course.png',bbox_inches='tight',pad_inches=0.05,facecolor='white')
