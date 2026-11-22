"""Reproducible synthetic figures for PowerPoint interoperability comparisons."""
from pathlib import Path
import argparse
import json
import subprocess
import sys
import numpy as np
import matplotlib as mpl
mpl.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

HERE=Path(__file__).resolve().parent
mpl.rcParams.update({'font.family':'Arial','font.size':10,'axes.titlesize':12,
    'axes.titlepad':12,'axes.labelsize':11,'axes.linewidth':1,'axes.spines.top':False,
    'axes.spines.right':False,'xtick.labelsize':9,'ytick.labelsize':9,
    'legend.fontsize':9,'legend.frameon':False,'svg.fonttype':'none',
    'mathtext.fontset':'custom','mathtext.rm':'Arial','mathtext.it':'Arial:italic',
    'mathtext.bf':'Arial:bold','savefig.facecolor':'white','svg.hashsalt':'editable-ppt-gallery'})
COLORS=['#0072B2','#B77811','#009E73']

def math_curves():
    fig,ax=plt.subplots(figsize=(5.4,3.6));fig.subplots_adjust(left=.19,bottom=.19,right=.95,top=.83)
    x=np.geomspace(.01,10,160)
    for j,a in enumerate((.3,1.,3.)):
        y=x**1.5*(1+a*x**-.417)
        ax.loglog(x,y,color=COLORS[j],lw=1.7,label=rf'$\alpha={a:g}$')
    ax.set(xlim=(.01,10),ylim=(1e-3,200),xlabel=r'$\delta/\delta_0$',ylabel=r'$F/F_0$')
    ax.set_title('Scaling curves');ax.legend(loc='upper left')
    ax.text(.96,.08,r'$F\propto\delta^{3/2}$',transform=ax.transAxes,ha='right',fontsize=11)
    return fig

def mesh_zoom():
    fig,axs=plt.subplots(1,2,figsize=(7.4,3.4));fig.subplots_adjust(left=.10,bottom=.19,right=.96,top=.82,wspace=.42)
    segments=[]
    def warp(x,y):return np.column_stack([x+.15*np.sin(y*.9),y-.55*np.exp(-x*x/3)*np.exp(y/4)])
    for x in np.linspace(0,10,31):
        y=np.linspace(-8,0,80);segments.append(warp(np.full_like(y,x),y))
    for y in np.linspace(-8,0,25):
        x=np.linspace(0,10,100);segments.append(warp(x,np.full_like(x,y)))
    x=np.linspace(0,10,200);surface=warp(x,np.zeros_like(x))
    for ax,title,xlim,ylim in zip(axs,['Full mesh','Contact region'],[(-.2,10.2),(-.1,2.8)],[(-8.2,.4),(-2.6,.4)]):
        ax.add_collection(LineCollection(segments,colors='#45545E',linewidths=.65))
        ax.plot(*surface.T,color=COLORS[1],lw=1.8)
        ax.set(xlim=xlim,ylim=ylim,xlabel=r'$r$ ($\mu$m)',ylabel=r'$z$ ($\mu$m)');ax.set_title(title)
    return fig

def color_field():
    fig,ax=plt.subplots(figsize=(5.4,3.6));fig.subplots_adjust(left=.18,bottom=.19,right=.79,top=.83)
    x=np.linspace(0,6,180);y=np.linspace(-4,0,120);xx,yy=np.meshgrid(x,y)
    field=np.exp(-((xx-.3)**2/7+yy**2/5))
    im=ax.imshow(field,extent=(0,6,-4,0),origin='lower',cmap='YlGnBu',vmin=0,vmax=1,aspect='auto',interpolation='bilinear')
    for depth in (0,-1,-2,-3):
        ax.plot(x,depth-.35*np.exp(-x*x/5),color='#333333',lw=.75)
    ax.set(xlabel=r'$r$ ($\mu$m)',ylabel=r'$z$ ($\mu$m)',xlim=(0,6),ylim=(-4,0));ax.set_title('Displacement field')
    ca=fig.add_axes([.84,.23,.025,.52]);cb=fig.colorbar(im,cax=ca,ticks=[0,.5,1]);cb.set_label(r'$|\mathbf{u}|/u_0$',labelpad=8)
    return fig

def errorbars():
    fig,ax=plt.subplots(figsize=(5.4,3.6));fig.subplots_adjust(left=.19,bottom=.19,right=.95,top=.83)
    t=np.arange(0,11)
    for j,tau in enumerate((1.6,3.2,5.4)):
        y=1-np.exp(-(t+.25)/tau);err=.025+.015*np.sin(t*.7+j)**2
        ax.errorbar(t,y,yerr=err,color=COLORS[j],lw=1.5,marker='o',ms=3.5,capsize=2.5,label=rf'$\tau={tau:g}$ s')
    ax.set(xlim=(-.5,10.5),ylim=(0,1.08),xlabel=r'Time, $t$ (s)',ylabel=r'Response, $A/A_0$')
    ax.set_title('Response with error bars');ax.legend(loc='lower right')
    return fig

CASES=[('01_math',math_curves),('02_mesh',mesh_zoom),('03_field',color_field),('04_errorbars',errorbars)]
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--output',type=Path,default=HERE/'artifacts');args=ap.parse_args()
    args.output.mkdir(parents=True,exist_ok=True);manifest=[]
    for name,fn in CASES:
        folder=args.output/name;folder.mkdir(exist_ok=True)
        if (folder/'native.pptx').exists():raise FileExistsError('Use a fresh output directory')
        fig=fn();fig.savefig(folder/'source.svg',metadata={'Date':None});fig.savefig(folder/'source.png',dpi=180)
        width,height=fig.get_size_inches()*72;plt.close(fig)
        subprocess.run([sys.executable,str(HERE/'matplotlib_scene.py'),str(folder/'source.svg'),str(folder)],check=True)
        subprocess.run([sys.executable,str(HERE.parents[1]/'scripts/native_pptx.py'),str(folder/'scene.json'),str(folder/'native.pptx')],check=True)
        manifest.append({'case':name,'width':float(width),'height':float(height),'svg':'source.svg','native':'native.pptx'})
    (args.output/'cases.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')

if __name__=='__main__':main()
