import json, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
fp="/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"; fb="/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
fm.fontManager.addfont(fp); fm.fontManager.addfont(fb)
plt.rcParams['font.family']=fm.FontProperties(fname=fp).get_name()
d=json.load(open('muni.json'))
F={f['properties']['code']:f for f in d['features']}
def polys(f):
    g=f['geometry']; cs=g['coordinates']
    return [p[0] for p in cs] if g['type']=='MultiPolygon' else [cs[0]]
def draw(ax,code,fc,ec,lw=1,z=1):
    for ring in polys(F[code]):
        xs=[c[0] for c in ring]; ys=[c[1] for c in ring]
        ax.fill(xs,ys,fc=fc,ec=ec,lw=lw,zorder=z)
NAVY="#0B1F4D"; OR="#C2410C"; G="#8C95A6"
# ---- inset: Daegu & Cheongdo
fig,ax=plt.subplots(figsize=(4.2,4.6),dpi=200)
for c in [k for k in F if k.startswith('22')]: draw(ax,c,"#E3E8F2","#FFFFFF",0.6)
for c in ['37100','37070','38080']: draw(ax,c,"#F3F4F6","#FFFFFF",0.6)
draw(ax,'37360',"#FFE3CC",OR,1.4,2)
ax.text(128.52,35.86,"대구",fontsize=13,weight='bold',color=NAVY,ha='center')
ax.text(128.80,35.83,"경산",fontsize=10,color=G,ha='center')
ax.text(128.88,35.97,"영천",fontsize=10,color=G,ha='center')
ax.text(128.70,35.45,"밀양",fontsize=10,color=G,ha='center')
ax.text(128.76,35.66,"청도군",fontsize=14,weight='bold',color=OR,ha='center')
ax.plot([128.6286],[35.8797],'o',ms=7,color=NAVY,zorder=5); ax.text(128.645,35.893,"동대구역",fontsize=9,color=NAVY)
ax.plot([128.7466],[35.6405],'s',ms=7,color=NAVY,zorder=5); ax.text(128.762,35.628,"청도역",fontsize=9,color=NAVY)
ax.annotate("",xy=(128.744,35.652),xytext=(128.632,35.868),arrowprops=dict(arrowstyle="-|>",color=NAVY,lw=1.6,ls=(0,(4,2))),zorder=4)
ax.text(128.715,35.775,"직선 약 29km\n기차 20~33분\n하루 약 20회",fontsize=8.5,color=NAVY,bbox=dict(fc="white",ec="none",pad=1.5))
ax.set_xlim(128.35,129.05); ax.set_ylim(35.42,36.02); ax.set_aspect(1/0.81); ax.axis('off')
plt.savefig('inset.png',bbox_inches='tight',pad_inches=0.05,transparent=False,facecolor='white')
# ---- detail
fig,ax=plt.subplots(figsize=(8.2,5.2),dpi=200)
draw(ax,'37360',"#FBF7F2","#C9B9A6",1.2)
for c in ['37100','37070','37020','38080','22310','38330' if '38330' in F else '37100']: 
    try: draw(ax,c,"#F4F5F7","#FFFFFF",0.8,0)
    except: pass
draw(ax,'37360',"#FBF7F2","#B9A58F",1.4,1)
P={ "청도역 · 상상마루\n(옛 버스터미널, 역 인근)":(128.7466,35.6405,"o",OR),
    "와인터널":(128.7205,35.7151,"o",OR),
    "소싸움경기장":(128.7233,35.6858,"o",OR),
    "신화랑풍류마을\n·오토캠핑장":(128.9223,35.7181,"o",OR),
    "운문사":(128.9566,35.6740,"o",OR),
    "대천터미널":(128.9143,35.7164,"D",NAVY),
    "동곡터미널":(128.8907,35.6934,"D",NAVY),
    "풍각정류장":(128.6154,35.6406,"D",NAVY)}
off={"청도역 · 상상마루\n(옛 버스터미널, 역 인근)":(0.009,-0.016),"와인터널":(0.008,0.006),"소싸움경기장":(-0.078,-0.004),
 "신화랑풍류마을\n·오토캠핑장":(0.01,0.008),"운문사":(0.008,-0.012),"대천터미널":(-0.055,-0.004),"동곡터미널":(-0.02,-0.016),"풍각정류장":(0.008,0.004)}
route=["청도역 · 상상마루\n(옛 버스터미널, 역 인근)","와인터널","소싸움경기장","신화랑풍류마을\n·오토캠핑장","운문사"]
import math
LOFF=[(0.022,-0.028),(-0.033,0.0),(0.0,-0.013),(0.012,0.002)]
for i,(a,b) in enumerate(zip(route,route[1:])):
    (x1,y1),(x2,y2)=P[a][:2],P[b][:2]
    km=math.hypot((x2-x1)*90.6,(y2-y1)*111)*1.35; mn=round(km/40*60)
    mx,my=(x1+x2)/2+LOFF[i][0],(y1+y2)/2+LOFF[i][1]
    ax.text(mx,my,f"약 {km:.0f}km · {mn}분",fontsize=8,color="#7A2E0B",ha='center',va='center',zorder=8,bbox=dict(boxstyle="round,pad=0.25",fc="#FFF4EC",ec="#F2C9A8",lw=0.8))
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),arrowprops=dict(arrowstyle="-|>",color=OR,lw=1.8,shrinkA=6,shrinkB=6,connectionstyle="arc3,rad=0.15"),zorder=3)
nums={r:i+1 for i,r in enumerate(route)}
for k,(x,y,m,c) in P.items():
    ax.plot([x],[y],m,ms=13 if c==OR else 7,color=c,mec='white',mew=1.2,zorder=5)
    if k in nums: ax.text(x,y,str(nums[k]),fontsize=7,color='white',weight='bold',ha='center',va='center',zorder=7)
    dx,dy=off[k]; ax.text(x+dx,y+dy,k,fontsize=8.5 if c!=OR else 9.5,weight='bold' if c==OR else 'normal',color=c if c!=G else "#5A6170",zorder=6)
ax.text(128.60,35.565,"이동: 택시 기준 추정 (직선거리×1.35, 시속 40km)",fontsize=7.5,color="#8C95A6",zorder=8)
ax.text(128.83,35.58,"청도군",fontsize=18,weight='bold',color="#D8C9B8",ha='center')
ax.set_xlim(128.57,129.06); ax.set_ylim(35.55,35.78); ax.set_aspect(1/0.81); ax.axis('off')
plt.savefig('detail.png',bbox_inches='tight',pad_inches=0.05,facecolor='white')
