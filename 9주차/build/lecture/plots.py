# 9·10주차 강의안 그래프 (matplotlib, Times 계열 수식)
import numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams.update({"font.family":"serif","font.serif":["Liberation Serif"],"mathtext.fontset":"stix","font.size":20,"axes.edgecolor":"#191919","axes.linewidth":1.2})
INK,BLUE,GRAY,NAVY,RED="#191919","#2563EB","#667085","#0B1F3A","#B42318"
def clean(ax):
    for s in ["top","right"]: ax.spines[s].set_visible(False)
# Legendre P0..P3
x=np.linspace(-1,1,400)
fig,ax=plt.subplots(figsize=(6.4,4.4),dpi=200)
for P,lab,c,ls in [(np.ones_like(x),r"$P_0$",GRAY,"--"),(x,r"$P_1$",INK,"-"),((3*x**2-1)/2,r"$P_2$",BLUE,"-"),((5*x**3-3*x)/2,r"$P_3$",NAVY,":")]:
    ax.plot(x,P,color=c,lw=3,ls=ls,label=lab)
ax.axhline(0,color="#D0D5DD",lw=1); ax.set_xlim(-1,1); ax.set_ylim(-1.1,1.15)
ax.set_xlabel(r"$x=\cos\alpha$"); ax.legend(frameon=False,ncol=4,loc="upper center",bbox_to_anchor=(0.5,1.15),fontsize=18); clean(ax)
fig.tight_layout(); fig.savefig("assets/legendre.png",transparent=True); plt.close(fig)
# dipole field lines: pure vs physical
def field_pure(X,Z):
    r2=X**2+Z**2; r5=r2**2.5
    return (3*X*Z)/r5, (3*Z*Z-r2)/r5
def field_phys(X,Z,d=0.5):
    def pc(x0,z0,q):
        dx,dz=X-x0,Z-z0; r3=(dx**2+dz**2)**1.5; return q*dx/r3,q*dz/r3
    a=pc(0,d/2,1); b=pc(0,-d/2,-1); return a[0]+b[0],a[1]+b[1]
g=np.linspace(-2,2,400); X,Z=np.meshgrid(g,g)
fig,axs=plt.subplots(1,2,figsize=(11,5.2),dpi=200)
for ax,(f,title) in zip(axs,[(field_pure,"순수 쌍극자"),(field_phys,"물리적 쌍극자")]):
    Ex,Ez=f(X,Z); mag=np.hypot(Ex,Ez)
    seeds=[(0.12*np.cos(t),0.12*np.sin(t)) for t in np.linspace(0.15,np.pi-0.15,9)] if f is field_pure else [(0.05*np.cos(t),0.25+0.05*np.sin(t)) for t in np.linspace(0,2*np.pi,16,endpoint=False)]
    ax.streamplot(X,Z,Ex,Ez,color=BLUE,linewidth=1.8,density=0.9,arrowsize=1.6)
    if f is field_pure:
        ax.annotate("",xy=(0,0.18),xytext=(0,-0.18),arrowprops=dict(arrowstyle="-|>",lw=3,color=INK))
    else:
        ax.plot([0],[0.25],"o",color=RED,ms=10); ax.plot([0],[-0.25],"o",color=INK,ms=10)
    ax.set_xlim(-2,2); ax.set_ylim(-2,2); ax.set_aspect("equal"); ax.axis("off")
    ax.set_title(title,fontsize=22,family="Noto Sans CJK KR",color=INK)
fig.tight_layout(); fig.savefig("assets/dipole_fields.png",transparent=True); plt.close(fig)
# uniformly polarized sphere E field (P along z): inside E=-P/3, outside dipole
def E_sphere(X,Z,R=1.0):
    r=np.hypot(X,Z); Ex=np.zeros_like(X); Ez=np.zeros_like(X)
    ins=r<R; Ez[ins]=-1/3
    p=R**3/3; r5=np.where(ins,1,r)**5; r2=np.where(ins,1,r)**2
    Ex[~ins]=(p*3*X*Z/r5)[~ins]; Ez[~ins]=(p*(3*Z*Z-r2)/r5)[~ins]
    return Ex,Ez
g=np.linspace(-2.4,2.4,500); X,Z=np.meshgrid(g,g); Ex,Ez=E_sphere(X,Z)
fig,ax=plt.subplots(figsize=(5.6,5.6),dpi=200)
ax.add_patch(plt.Circle((0,0),1,color="#E0F0FE",zorder=0)); ax.add_patch(plt.Circle((0,0),1,fill=False,color=NAVY,lw=2))
ax.streamplot(X,Z,Ex,Ez,color=BLUE,linewidth=1.6,density=0.9,arrowsize=1.5)
ax.set_xlim(-2.4,2.4); ax.set_ylim(-2.4,2.4); ax.set_aspect("equal"); ax.axis("off")
fig.tight_layout(); fig.savefig("assets/polarized_sphere_E.png",transparent=True); plt.close(fig)
print("ok")
