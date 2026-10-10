/* 場景動畫（倉庫貨場：棚子、貨櫃、工人、貓狗、蝴蝶、天色與天氣）
 * 取自同作者的 UPS Reprice Platform（sandyliu3056/ups-reprice-web，index.html 的 SCENE 區塊），
 * 原樣搬過來，只把語言／時區／招牌文字改成由 RMAScene.init() 設定。
 * 天色依「時區下拉」的當地時間變化（白天太陽、夜裡月亮與星星）；點角色會有反應。 */
(function () {
  "use strict";
  const $ = (s) => document.querySelector(s);
  const BRAND = "Tally";
  const SC = { lang: "zh", tz: "Asia/Taipei", sign: "RMA" };
  function pal(k, f) {
    const v = getComputedStyle(document.documentElement).getPropertyValue(k).trim();
    return v || f;
  }
  function tzParts() {
    const now = new Date(), p = {};
    let z = SC.tz;
    try { new Intl.DateTimeFormat("en-US", { timeZone: z }); } catch (e) { z = "Asia/Taipei"; }
    for (const x of new Intl.DateTimeFormat("en-US", { timeZone: z, hour12: false, hour: "2-digit", minute: "2-digit", second: "2-digit",
        year: "numeric", month: "2-digit", day: "2-digit", weekday: "short" }).formatToParts(now)) p[x.type] = x.value;
    return { h: +p.hour % 12, H: +p.hour % 24, m: +p.minute, s: +p.second, y: +p.year, mo: +p.month, d: +p.day, wd: p.weekday };
  }

function noise(a,b){
  let n=((a|0)*374761393 + (b|0)*668265263) & 0xFFFFFFFF;
  n=((n ^ (n>>>13)) * 1274126177) & 0xFFFFFFFF;
  return ((n ^ (n>>>16)) & 0xFFFF)/32767.5 - 1.0;
}
const SCENE={cv:null,ctx:null,t:0,actors:[],raf:null,run:0,W:0,H:0,dpr:1,progress:0};

function skLine(g,x1,y1,x2,y2,k){
  const wob=1.1, n1=noise(x1*7+k,y1*3), n2=noise(x2*5+k,y2*11);
  g.beginPath(); g.moveTo(x1,y1);
  g.quadraticCurveTo((x1+x2)/2+n1*wob*2,(y1+y2)/2+n2*wob*2,x2,y2);
  g.stroke();
}
function skRect(g,x,y,w,h,k,fill){
  const p=[[x,y],[x+w,y],[x+w,y+h],[x,y+h]].map(([px,py],i)=>
    [px+noise(px+k,py+i)*1.2, py+noise(py+k,px+i)*1.2]);
  g.beginPath(); g.moveTo(p[0][0],p[0][1]);
  for(let i=1;i<4;i++) g.lineTo(p[i][0],p[i][1]);
  g.closePath();
  if(fill){ g.fillStyle=fill; g.fill(); }
  g.stroke();
}
function skCircle(g,cx,cy,r,k,fill){
  g.beginPath();
  for(let a=0;a<=Math.PI*2+0.01;a+=Math.PI/9){
    const rr=r+noise(cx+k+a*40,cy)*0.9;
    const px=cx+Math.cos(a)*rr, py=cy+Math.sin(a)*rr;
    a===0?g.moveTo(px,py):g.lineTo(px,py);
  }
  g.closePath();
  if(fill){ g.fillStyle=fill; g.fill(); }
  g.stroke();
}

/* ---------------------------------------------------------------------
   角色
   ---------------------------------------------------------------------
   數字照桌面版 _doodle_person()搬,不是憑印象畫:
     頭  hy = 地板 - 51,頭塊 x±14.5、hy-14..hy+13(29 寬 27 高)
     身  y-37 .. y-19(18 高)   腿 y-19 .. y-6   鞋 15 寬 6.6 高往外攤
   頭幾乎佔全高四成。照真人比例畫(小頭長腿)就會變成別的畫風。

   兩件不能省的細節,原始碼都特別註明過:
     頰上三道短線 —— 拿掉就少了這個畫法的味道
     背心的開襟和反光帶 —— 不留開襟就只是一塊色斑,看不出是背心

   貓狗毛色是奶油白加色塊(CAT_COAT #f2ece1 / DOG_COAT #f5efe6),
   靠輪廓分:貓尖耳細尾坐姿偏高,狗垂耳粗尾上翹身上有色塊。
   --------------------------------------------------------------------- */
const BLUSH="#f0a9a4", SKIN="#ffe0bf";
function shadow(g,x,y,w){
  g.save(); g.globalAlpha=.14; g.fillStyle=INK;
  g.beginPath(); g.ellipse(x,y+1,w,3.2,0,0,7); g.fill(); g.restore();
}
function blob(g,x1,y1,x2,y2,fill,r){
  const w=x2-x1,h=y2-y1; r=Math.min(r,w/2,h/2);
  g.beginPath();
  g.moveTo(x1+r,y1);
  g.arcTo(x2,y1,x2,y2,r); g.arcTo(x2,y2,x1,y2,r);
  g.arcTo(x1,y2,x1,y1,r); g.arcTo(x1,y1,x2,y1,r);
  g.closePath(); g.fillStyle=fill; g.fill(); g.stroke();
}
function limb(g,x1,y1,x2,y2,color,w){
  g.save(); g.lineCap="round";
  g.strokeStyle=INK; g.lineWidth=w+2.2;
  g.beginPath(); g.moveTo(x1,y1); g.lineTo(x2,y2); g.stroke();
  g.strokeStyle=color; g.lineWidth=w;
  g.beginPath(); g.moveTo(x1,y1); g.lineTo(x2,y2); g.stroke();
  g.restore();
}
function dot(g,x,y,rx,ry){ g.fillStyle=INK;
  g.beginPath(); g.ellipse(x,y,rx,ry,0,0,7); g.fill(); }
function patch(g,x,y,rx,ry,c){ g.save(); g.fillStyle=c;
  g.beginPath(); g.ellipse(x,y,rx,ry,0,0,7); g.fill(); g.restore(); }
/* 頰上三道短線 */
function cheekLines(g,x,hy,sd){
  g.strokeStyle=INK; g.lineWidth=1.1;
  for(let k=0;k<3;k++){ const ly=hy+0.4+k*2.0;
    g.beginPath(); g.moveTo(x+sd*12.4,ly); g.lineTo(x+sd*8.4,ly+1.1); g.stroke(); }
}
function hearts(g,a,t){
  if(!a.love) return;
  for(let i=0;i<3;i++){
    const p=(t-a.love+i*8)/38; if(p<0||p>1) continue;
    g.save(); g.globalAlpha=1-p; g.fillStyle="#e05a6a";
    const hx=a.x+(i-1)*9, hy=a.y-64-p*26, s=4;
    g.beginPath(); g.moveTo(hx,hy+s);
    g.bezierCurveTo(hx-s*1.6,hy-s*.4,hx-s*.5,hy-s*1.6,hx,hy-s*.4);
    g.bezierCurveTo(hx+s*.5,hy-s*1.6,hx+s*1.6,hy-s*.4,hx,hy+s);
    g.fill(); g.restore();
  }
  if(t-a.love>70) a.love=0;
}

/* 戶外而且在下雨或下雪才要雨具。Scoobi 是倉庫裡面,不用穿。
   WX 與 BRAND 都在後面才宣告,但這個函式只在繪圖時被呼叫,
   那時兩個都已經有值了。 */
function outdoorWet(){
  return typeof BRAND!=="undefined" && BRAND==="Tally"
      && typeof WX!=="undefined"
      && (WX.mode==="rain"||WX.mode==="snow"||WX.mode==="storm");
}

function drawWorker(g,a,t){
  const x=a.x, y=a.y;
  const sw=a.act==="walk"?Math.sin(t/8+a.seed)*2.6:0;
  const lift=a.act==="lift"?6:0;
  const blink=(Math.floor(t/7)+a.seed)%48===0;
  shadow(g,x,y,17);
  g.strokeStyle=INK; g.lineWidth=2.1; g.lineJoin="round";

  /* 腿短、鞋寬且往外攤 —— 鞋是最容易畫小的一塊 */
  for(const sd of [-1,1]){
    const lx=x+sd*5.4+sw*sd*.45;
    limb(g,lx,y-19,lx+sd*.5,y-6,"#4a5560",7);
    g.lineWidth=2.1;
    blob(g,lx-8+sd*1.6,y-6.6,lx+7+sd*1.6,y,"#2f2a26",3);
  }
  /* 手臂先畫,身體蓋上去;手掌等身體畫完再補,不然會被蓋住。 */
  const hands=[];
  const waving=a.act==="wave";
  const umbrella=outdoorWet();
  for(const sd of [-1,1]){
    /* 揮手:手腕左右擺,不是舉著不動 —— 靜態的舉手看起來像在攔車。 */
    let tipx=x+sd*15+sd*sw*.35, tipy=y-22-lift;
    if(waving && sd>0){
      const wob=Math.sin(t/4)*5.4;
      tipx=x+19+wob; tipy=y-48-Math.abs(wob)*0.22;
    }
    /* 撐傘那隻手舉到頭上,另一隻自然垂著 */
    if(umbrella && sd>0){ tipx=x+7; tipy=y-62; }
    limb(g,x+sd*10,y-34,tipx,tipy,a.coat,5.8);
    hands.push([tipx,tipy]);
  }
  if(umbrella){
    const ux=x+7, uy=y-62;
    g.strokeStyle=INK; g.lineWidth=2;
    skLine(g,ux,uy,ux,uy-16,a.seed+40);          /* 傘柄 */
    const canopy=pal("--tab","#ffb500");
    g.fillStyle=canopy;
    g.beginPath();
    g.moveTo(ux-24,uy-16);
    g.quadraticCurveTo(ux,uy-38,ux+24,uy-16);
    /* 傘面下緣是三個弧,不是一條直線 —— 直線看起來像帽子 */
    g.quadraticCurveTo(ux+16,uy-12,ux+8,uy-16);
    g.quadraticCurveTo(ux,uy-12,ux-8,uy-16);
    g.quadraticCurveTo(ux-16,uy-12,ux-24,uy-16);
    g.closePath(); g.fill(); g.stroke();
    g.lineWidth=1.2;
    skLine(g,ux,uy-34,ux-8,uy-16,a.seed+41);
    skLine(g,ux,uy-34,ux+8,uy-16,a.seed+42);
    g.lineWidth=2;
    skLine(g,ux,uy-38,ux,uy-42,a.seed+43);       /* 頂端小突起 */
  }
  g.lineWidth=2.1;
  /* 身體:矮而寬 */
  blob(g,x-11.5,y-37,x+11.5,y-19,a.coat,5);
  if(a.vest){
    blob(g,x-11,y-35.5,x+11,y-20,a.vest,2.8);
    g.strokeStyle=a.coat; g.lineWidth=2.4;              /* 開襟 */
    g.beginPath(); g.moveTo(x,y-35.5); g.lineTo(x,y-20); g.stroke();
    g.strokeStyle="#f6f1e6"; g.lineWidth=2.8;           /* 反光帶 */
    g.beginPath(); g.moveTo(x-10.6,y-28); g.lineTo(x+10.6,y-28); g.stroke();
    g.strokeStyle=INK; g.lineWidth=2.1;
  }
  for(const [hx2,hy2] of hands) blob(g,hx2-4.6,hy2-4,hx2+4.6,hy2+5,SKIN,4);
  if(a.box){ blob(g,x-11,y-50,x+11,y-36,BOX,2.5);
    g.beginPath(); g.moveTo(x-11,y-43); g.lineTo(x+11,y-43); g.stroke(); }

  /* 頭:比身體寬,幾乎佔全高四成 */
  const hy=y-51;
  for(const sd of [-1,1]) blob(g,x+sd*14-3.2,hy+1,x+sd*14+3.2,hy+8,SKIN,2.5);
  blob(g,x-14.5,hy-14,x+14.5,hy+13,SKIN,6.5);
  /* 安全帽:半圓 + 帽簷 */
  g.lineWidth=2.4; g.fillStyle=a.hat;
  g.beginPath(); g.ellipse(x,hy-10,14,14,0,Math.PI,0); g.closePath();
  g.fill(); g.stroke();
  blob(g,x-17.5,hy-13,x+17.5,hy-6,a.hat,3.4);
  g.lineWidth=2.1;
  for(const sd of [-1,1]){
    if(blink){ g.lineWidth=1.9;
      g.beginPath(); g.moveTo(x+sd*5.6-2.1,hy-1.4);
      g.lineTo(x+sd*5.6+2.1,hy-1.4); g.stroke(); g.lineWidth=2.1; }
    else dot(g,x+sd*5.6,hy-1.4,2.1,2.2);
    patch(g,x+sd*10,hy+3.3,3.4,1.9,BLUSH);
    cheekLines(g,x,hy,sd);
  }
  g.strokeStyle=INK; g.lineWidth=1.9; g.lineCap="round";
  g.beginPath(); g.ellipse(x,hy+3.4,4.4,3.6,0,
    18*Math.PI/180,162*Math.PI/180); g.stroke();   /* 微笑 */
  hearts(g,a,t);
}

/* 狗:垂耳、粗尾上翹、身上有色塊 */
/* 雨衣:蓋在身體上的一層,加一個帽兜。顏色用主題的金色,
   和工人的傘同一組,看起來才像同一套裝備。 */
function coatOver(g,x,y,w,hh,hx,hy,hr,seed){
  const c=pal("--tab","#ffb500");
  g.strokeStyle=INK; g.lineWidth=1.9; g.fillStyle=c;
  g.beginPath();
  g.moveTo(x-w,y);
  g.quadraticCurveTo(x-w-1,y-hh*0.7, x-w*0.5,y-hh);
  g.quadraticCurveTo(x,y-hh-2, x+w*0.5,y-hh);
  g.quadraticCurveTo(x+w+1,y-hh*0.7, x+w,y);
  g.closePath(); g.fill(); g.stroke();
  g.lineWidth=1.1;
  skLine(g,x,y-hh*0.85,x,y-2,seed+60);
  g.lineWidth=1.9;
  /* 帽兜 */
  g.fillStyle=c;
  g.beginPath();
  g.arc(hx,hy,hr+3.2,Math.PI*0.98,Math.PI*2.02);
  g.closePath(); g.fill(); g.stroke();
}

function drawDog(g,a,t){
  /* 蝴蝶飛近:原地小跳、耳朵豎起、尾巴搖得飛快;停在頭上:抬眼看,最後甩甩頭 */
  const pose=dogPose(a,t);
  const x=a.x, y=a.y-pose.hop, f=a.dir<0?-1:1, wag=Math.sin(t/(pose.excited?1.6:3.6))*6;
  const wet=outdoorWet();
  const blink=(Math.floor(t/7)+a.seed)%52===0 && !pose.sit;
  shadow(g,x,a.y,20);
  g.strokeStyle=INK; g.lineWidth=2.1; g.lineJoin="round";
  /* 粗尾上翹 */
  limb(g,x-15*f,y-20,x-24*f,y-31+wag,DOG_COAT,7);
  g.lineWidth=2.1;
  const st=a.act==="walk"?Math.sin(t/9)*3:0;
  const pee=a.act==="pee";
  for(const dx of [-11,-4,5,12]){
    /* 尿尿時後腿抬起來 */
    const up=(pee&&dx===-11)?8:0;
    const lx=x+dx*f+st*(dx<0?1:-1);
    limb(g,lx,y-13-up,lx,y-3-up,DOG_COAT,5.4);
    g.lineWidth=1.8;
    blob(g,lx-4,y-4.5-up,lx+4,y-up,DOG_COAT,2.6);
  }
  if(pee){
    g.save(); g.strokeStyle="#e8c65a"; g.lineWidth=1.6; g.lineCap="round";
    const n=(t%18)/18;
    g.beginPath();
    g.moveTo(x-13*f,y-9);
    g.quadraticCurveTo(x-18*f,y-6,x-21*f-n*3,y-1);
    g.stroke();
    g.globalAlpha=.55; g.fillStyle="#e8c65a";
    g.beginPath(); g.ellipse(x-23*f,y,5+n*2,1.8,0,0,7); g.fill();
    g.restore();
    g.strokeStyle=INK; g.lineWidth=2.1;
  }
  blob(g,x-16,y-25,x+16,y-6,DOG_COAT,8);
  patch(g,x-6*f,y-19,7,5,DOG_MARK);                    /* 身上的色塊 */
  const hx=x+17*f+pose.shake, hy=y-36, earUp=pose.excited?5:0;
  /* 雨衣蓋在身體上、頭之前畫,帽兜才會壓在耳朵後面 */
  if(wet) coatOver(g,x,y-4,17,23,hx,hy,13,a.seed);
  blob(g,hx-13.5,hy-12,hx+13.5,hy+12,DOG_COAT,7);
  for(const sd of [-1,1])                              /* 垂耳;興奮時豎起來 */
    blob(g,hx+sd*13-5,hy-9-earUp,hx+sd*13+5,hy+7-earUp,DOG_MARK,4.5);
  patch(g,hx+5*f,hy+4,8,5.5,DOG_MARK);
  for(const sd of [-1,1]){
    if(blink){ g.lineWidth=1.9; g.beginPath();
      g.moveTo(hx+sd*5.4-2,hy-2); g.lineTo(hx+sd*5.4+2,hy-2); g.stroke(); g.lineWidth=2.1; }
    else dot(g,hx+sd*5.4,hy-2-(pose.sit?1.6:0),2.1,2.2);   /* 停著就往上看 */
    patch(g,hx+sd*10,hy+3,3.2,1.8,BLUSH);
  }
  dot(g,hx,hy+4.4,2.6,2);                              /* 鼻 */
  if(a.act==="bark"){
    g.fillStyle="#e0707a"; g.beginPath();
    g.ellipse(hx,hy+9.5,2.8,3.6,0,0,7); g.fill(); g.stroke();
    g.font='700 12px "Comic Sans MS",system-ui'; g.fillStyle=INK;
    g.fillText((SC.lang==="en")?"woof!":"汪!",hx+15*f,hy-14);
  }
  hearts(g,a,t);
}

/* 貓:尖耳、細尾、坐姿。身體是下寬上窄貼在地上的一團,
   不是兩隻腳站著 —— 站姿會看成別的動物。 */
function drawCat(g,a,t){
  const walking=a.act==="walk";
  const step=walking?Math.sin(t/5)*3:0;
  const bob=walking?Math.abs(Math.sin(t/5))*1.6:0;
  const x=a.x, y=a.y-bob, sway=Math.sin(t/7)*4;
  const blink=(Math.floor(t/7)+a.seed)%44===0;
  const puff=a.act==="puff";
  const wet=outdoorWet();
  const hfPre=a.dir<0?-1:1;   /* 面向哪邊,尾巴接在相反端 */
  /* 蝴蝶停在頭上:醒著、往上盯、耳朵抖;最後 45 幀眼睛瞪大。手不動 —— 舉手拍看起來很怪。 */
  const bug=a.bug&&a.bug.act==="sit"?a.bug:null, stare=!!(bug&&bug.wait<45&&!walking);
  const sleep=a.act==="sit"&&!puff&&!bug&&(Math.floor(t/260)%2===1);
  shadow(g,x,y,puff?18:15);
  g.strokeStyle=INK; g.lineWidth=2.1; g.lineJoin="round";

  /* 尾巴:炸毛時豎起來變粗 */
  if(puff){
    limb(g,x+10,y-14,x+16,y-46,CAT_COAT,9);
    g.strokeStyle=INK; g.lineWidth=1.4;
    for(const [tx,ty] of [[13,-24],[18,-33],[12,-40]]){
      g.beginPath(); g.moveTo(x+tx,y+ty); g.lineTo(x+tx+7,y+ty-5); g.stroke();
    }
  }else{
    limb(g,x-(walking?12*hfPre:-11),y-(walking?16:14),
         x-(walking?20*hfPre:-21),y-(walking?30:30)+sway,CAT_COAT,4.2);
  }
  g.strokeStyle=INK; g.lineWidth=2.1;

  const bw=puff?17:13, bh=puff?26:24;
  if(walking){
    /* 走路是四腳站姿:身體橫躺,四條腿在下面交替。
       坐姿那團下寬上窄的形狀走起來會像在滑行。 */
    const f=a.dir<0?-1:1;
    for(const dx of [-9,-3,4,10]){
      const sw2=Math.sin(t/5 + (dx<0?0:Math.PI))*3;
      const lx=x+dx*f+sw2;
      limb(g,lx,y-13,lx,y-3,CAT_COAT,4.4);
      g.lineWidth=1.8;
      g.fillStyle=CAT_COAT;
      g.beginPath(); g.ellipse(lx,y-2,3.4,2.2,0,0,7); g.fill(); g.stroke();
      g.lineWidth=2.1;
    }
    blob(g,x-13,y-22,x+13,y-9,CAT_COAT,7);
    for(let i=0;i<2;i++) patch(g,x-4+i*7,y-18,2.4,3.2,CAT_MARK);
  }else{
    /* 坐姿身體:下緣貼地、上窄 */
    g.fillStyle=CAT_COAT;
    g.beginPath();
    g.moveTo(x-bw,y);
    g.quadraticCurveTo(x-bw-1,y-bh*0.62, x-bw*0.52,y-bh);
    g.quadraticCurveTo(x,y-bh-3, x+bw*0.52,y-bh);
    g.quadraticCurveTo(x+bw+1,y-bh*0.62, x+bw,y);
    g.closePath(); g.fill(); g.stroke();
    /* 前腳:貼在身體前緣的兩顆 */
    for(const dx of [-5,4]){
      g.fillStyle=CAT_COAT;
      g.beginPath(); g.ellipse(x+dx,y-2.5,4.2,2.8,0,0,7); g.fill(); g.stroke();
    }
  }
  for(let i=0;i<2;i++) patch(g,x-2+i*6,y-16,2,4,CAT_MARK);
  /* 炸毛:身體周圍豎起的毛。穿著雨衣時不畫 —— 毛會穿出衣服。 */
  if(puff && !wet){
    g.strokeStyle=INK; g.lineWidth=1.5;
    for(let k=0;k<9;k++){
      const ang=Math.PI+k*Math.PI/8;
      const px=x+Math.cos(ang)*bw, py=y-bh*0.5+Math.sin(ang)*bh*0.5;
      g.beginPath(); g.moveTo(px,py);
      g.lineTo(px+Math.cos(ang)*7,py+Math.sin(ang)*7); g.stroke();
    }
    g.lineWidth=2.1;
  }

  /* 走路時頭低而且在身體前端,坐著才是高高在上 */
  const hf=a.dir<0?-1:1;
  const hxo=walking?14*hf:0;
  const hy=walking ? y-24 : y-(puff?38:34);
  /* 雨衣蓋身體、帽兜罩頭,要等 hxo/hy 算完才知道頭在哪 */
  /* 炸毛的時候雨衣不會消失,只是被撐開一點。 */
  if(wet) coatOver(g,x, y-(walking?7:1),
    walking?14:(puff?bw+4:bw+2), walking?16:bh-2,
    x+hxo, hy, puff?15:13, a.seed);
  for(const sd of [-1,1]){
    const flick=(bug&&sd===hf&&(t%70)<10)?Math.sin(t/1.5)*2:0;   /* 抖耳朵 */
    g.fillStyle=CAT_COAT; g.beginPath();
    g.moveTo(x+sd*4,hy-8); g.lineTo(x+hxo+sd*10+flick,hy-22); g.lineTo(x+hxo+sd*13,hy-4);
    g.closePath(); g.fill(); g.stroke();
    patch(g,x+hxo+sd*9,hy-11,2.6,4.4,BLUSH);
  }
  blob(g,x+hxo-11,hy-10,x+hxo+11,hy+10,CAT_COAT,7);
  for(const sd of [-1,1]){
    if(puff||stare){                           /* 炸毛、盯蝴蝶時眼睛瞪大 */
      g.fillStyle="#fff";
      g.beginPath(); g.ellipse(x+hxo+sd*4.6,hy-1,3.4,3.8,0,0,7); g.fill(); g.stroke();
      dot(g,x+hxo+sd*4.6,hy-1-(stare?1.4:0),1.5,2.2);
    }
    else if(blink||sleep){ g.lineWidth=1.9; g.beginPath();
      g.moveTo(x+sd*5.4-2.2,hy-1); g.lineTo(x+sd*5.4+2.2,hy-1); g.stroke(); g.lineWidth=2.1; }
    else dot(g,x+hxo+sd*4.6,hy-1-(bug?2:0),2.1,2.3);   /* 停著就往上盯 */
    patch(g,x+hxo+sd*8.4,hy+4,3.2,1.8,BLUSH);
    cheekLines(g,x+hxo,hy,sd);
    g.strokeStyle=INK; g.lineWidth=.9;
    for(const dy of [-.5,2.2]){ g.beginPath();
      g.moveTo(x+hxo+sd*9.6,hy+2+dy); g.lineTo(x+hxo+sd*17,hy+dy*1.7); g.stroke(); }
    g.lineWidth=2.1;
  }
  g.fillStyle="#e0707a"; g.beginPath();
  g.moveTo(x+hxo,hy+3.4); g.lineTo(x+hxo-2.2,hy+1.4); g.lineTo(x+hxo+2.2,hy+1.4);
  g.closePath(); g.fill();
  if(puff){                                    /* 張嘴哈氣 */
    g.fillStyle=INK; g.beginPath();
    g.ellipse(x+hxo,hy+6.5,3,2.4,0,0,7); g.fill();
    g.font='700 12px "Comic Sans MS",system-ui'; g.fillStyle=INK;
    g.fillText((SC.lang==="en")?"hiss!":"哈!",x+hxo+16,hy-16);
  }
  if(sleep){ g.font='700 12px system-ui'; g.fillStyle=INK;
    g.fillText("z",x+hxo+16,hy-14); g.fillText("z",x+hxo+21,hy-21); }
  hearts(g,a,t);
}

let INK="#351c15", BOX="#8a6104", PAPER="#fffdf7",
    CAT_COAT="#f2ece1", CAT_MARK="#cfc7b8",
    DOG_COAT="#f5efe6", DOG_MARK="#b5793f";
/* 顯示場景的唯一入口。三個地方原本各自寫 style.display,
   顯示之後都沒有重量尺寸 —— 那就是動畫不見的原因。 */
function sceneShow(on){
  const sb=$("#sceneBox");
  if(!sb) return;
  sb.style.display = on?"":"none";
  if(on && typeof sceneResize==="function" && SCENE && SCENE.cv){
    sceneResize();
    if(typeof spreadActors==="function") spreadActors();
    if(typeof sceneDraw==="function") sceneDraw();
  }
}
function sceneResize(){
  const cv=SCENE.cv; if(!cv) return;
  const r=cv.getBoundingClientRect();
  SCENE.dpr=window.devicePixelRatio||1;
  cv.width=r.width*SCENE.dpr; cv.height=r.height*SCENE.dpr;
  SCENE.W=r.width; SCENE.H=r.height;
  SCENE.ctx.setTransform(SCENE.dpr,0,0,SCENE.dpr,0,0);
}
function sceneInit(){
  SCENE.cv=$("#scene"); if(!SCENE.cv) return;
  SCENE.ctx=SCENE.cv.getContext("2d");
  const cs=getComputedStyle(document.documentElement);
  INK=cs.getPropertyValue("--accent").trim()||INK;
  sceneResize(); addEventListener("resize",()=>{ sceneResize(); spreadActors(); });
  SCENE.actors=[
    /* 襯衫和背心必須是對比色 —— 同色系疊起來只會看成衣服上一塊補丁 */
    {kind:"worker",x:72, y:150,seed:11,act:"walk",dir:1, box:false,tgt:230,
     coat:"#5b86b8",vest:"#e8762c",hat:"#ffb500",hair:"#3a2b20",love:0},
    {kind:"worker",x:214,y:156,seed:29,act:"idle",dir:-1,box:true, tgt:92,
     coat:"#d8d2c6",vest:"#5b9e6b",hat:"#e8762c",hair:"#6b4a2e",love:0},
    {kind:"dog",   x:150,y:162,seed:53,act:"walk",dir:1, tgt:300,love:0},
    {kind:"cat",   x:334,y:160,seed:71,act:"sit",love:0},
  ];
  spreadActors();
  bugsInit();
  /* 摸摸。桌面版的場景可以點角色互動( _pet_at),這裡同樣。 */
  SCENE.cv.style.cursor="pointer";
  SCENE.cv.title = (SC.lang==="en")?"Click an animal":"點一下摸摸";
  SCENE.cv.addEventListener("click",e=>{
    const r=SCENE.cv.getBoundingClientRect();
    const mx=e.clientX-r.left, my=e.clientY-r.top;
    let best=null,bd=1e9;
    for(const a of SCENE.actors){
      const d=Math.hypot(a.x-mx,(a.y-20)-my);
      if(d<bd){bd=d;best=a;}
    }
    if(best&&bd<42){
      best.love=SCENE.t;
      /* 點下去各有各的反應:狗抬腿尿尿,貓炸毛,工人揮手。 */
      if(best.kind==="dog"){ best.act="pee";  best.wait=90; }
      else if(best.kind==="cat"){ best.act="puff"; best.wait=110; }
      else { best.act="wave"; best.wait=70; }
    }
  });
  if(!matchMedia("(prefers-reduced-motion: reduce)").matches) sceneLoop();
  else sceneDraw();
}
/* 角色沿著整條場景散開,不要全擠在左邊。
   Tally 左三分之一是棚子和貨櫃,所以人往右擺;Scoobi 左邊是貨架。 */
function spreadActors(){
  const W=SCENE.W||900;
  const lead = BRAND==="Tally" ? Math.min(0.42*W, 420) : Math.min(0.18*W, 190);
  const span = Math.max(160, W-lead-70);
  const a=SCENE.actors;
  if(!a.length) return;
  /* 平均分配時貓是最後一個,會落在最右邊 —— Scoobi 那裡是堆高機
     和高架,貓擠在角落。把貓抽出來擺中間,其餘的平均分。 */
  const cat = BRAND==="Scoobi" ? a.find(x=>x.kind==="cat") : null;
  const rest = cat ? a.filter(x=>x!==cat) : a;
  rest.forEach((act,i)=>{
    act.x = lead + span*(i+0.5)/rest.length;
    act.home = act.x;                    /* 走動的起點,不會離太遠 */
    act.tgt = Math.min(W-50, act.x + span/rest.length*0.8);
    act.dir = 1;
    if(act.wait===undefined) act.wait = 120+((i*97)%240);
  });
  if(cat){
    /* 中間偏左一點,不然會跟平均分配到中點的工人疊在一起 */
    cat.x = lead + span*0.40;
    cat.home = cat.x;
    cat.tgt = Math.min(W-50, cat.x + span*0.14);
    cat.dir = 1;
    if(cat.wait===undefined) cat.wait = 200;
  }
}

function sceneStep(){
  const W=SCENE.W||640, FLR=(SCENE.H||190)-22;
  for(const a of SCENE.actors) a.y=FLR;
  if(SCENE.t%3) return;
  for(const a of SCENE.actors){
    if(a.kind==="cat"){
      /* 打雷會嚇到。雷是全場同一瞬間,所以讀 FLASH 而不是自己擲骰。 */
      /* 閃電那一幀就要嚇到 —— 慢一幀看起來像先聽到雷才反應。 */
      if(WX.mode==="storm" && SCENE.t===FLASH && a.act!=="puff"){
        a.act="puff"; a.wait=150; continue;
      }
      /* 炸毛是被點出來或被雷嚇的,時間到就坐回去 */
      if(a.act==="puff"){ if(--a.wait<=0) a.act="sit"; continue; }
      /* 貓也會走動:走一段、坐一下、再走。不是一直坐著。 */
      if(a.act==="walk"){
        a.x += 0.42*a.dir;
        if(a.x>a.tgt || a.x<a.home-10){ a.dir*=-1; a.x+=0.42*a.dir; }
        if(--a.wait<=0){ a.act="sit"; a.wait=240+((a.seed*37)%300); }
      }else{
        if(--a.wait<=0){ a.act="walk"; a.dir=Math.random()<0.5?-1:1;
                         a.wait=180+((a.seed*53)%260); }
      }
      continue;
    }
    if(a.kind==="worker" && a.act==="wave"){
      /* 一邊走一邊揮,停下來會看起來像當機。 */
      a.x += 0.55*a.dir;
      if(a.x>a.tgt || a.x<a.home-10){ a.dir*=-1; a.x+=0.55*a.dir; }
      if(--a.wait<=0) a.act="walk";
      continue;
    }
    if(a.act==="walk"){
      a.x += a.dir*(a.kind==="dog"?0.9:0.6);
      if((a.dir>0&&a.x>=a.tgt)||(a.dir<0&&a.x<=a.tgt)){
        a.dir*=-1;
        a.tgt = a.dir>0? Math.min(W-50,a.x+170) : Math.max(50,a.x-170);
        if(a.kind==="worker"){ a.act="lift"; a.wait=40; }
        else {
          const r=Math.random();
          if(r<0.35){ a.act="pee";  a.wait=55; }
          else if(r<0.6){ a.act="bark"; a.wait=30; }
          else { a.act="idle"; a.wait=25; }
        }
      }
    }else{
      if(--a.wait<=0){
        if(a.kind==="worker") a.box=!a.box;
        a.act="walk";
      }
    }
    a.x=Math.max(40,Math.min(W-40,a.x));
  }
}
/* Tally:招牌寫著 TALLY 的手繪工棚,天空、太陽、遠山、雲、鳥。
   Scoobi:室內倉庫,貨架與輸送帶,輸送帶跑報表時就是進度條。
   兩套是不同的場景,不是同一張圖換濾鏡。 */

/* =====================================================================
   天色與天氣
   ---------------------------------------------------------------------
   時間完全離線:時區下拉已經知道當地幾點,天色照那個小時算。
   天氣預設是手動選的,要即時天氣才會連網 —— 這個檔案其他地方
   一個網路請求都沒有,那一個要由使用者自己開。
   ===================================================================== */
const TZ_COORD=[[25.03,121.57],[40.71,-74.01],[41.88,-87.63],
                [39.74,-104.99],[34.05,-118.24],[51.48,-0.01]];
const WX_KEY="rma_wx";
let WX={mode:"clear", live:true, temp:null, at:0};
try{ const s=localStorage.getItem(WX_KEY); if(s) WX=Object.assign(WX,JSON.parse(s)); }catch(e){}
function wxSave(){ try{ localStorage.setItem(WX_KEY,JSON.stringify(WX)); }catch(e){} }

/* 當地幾點,回 0..24 的小數 */
function localHour(){
  try{ const t=tzParts(); return t.H + t.m/60; }catch(e){ return 12; }
}
/* 這個小時屬於哪一段天色 */
function skyPhase(h){
  if(h<5)  return "night";
  if(h<7)  return "dawn";
  if(h<17) return "day";
  if(h<19) return "dusk";
  if(h<21) return "evening";
  return "night";
}
const SKY={
  /* top 是天頂,bot 是地平線。晨昏的暖色只出現在地平線那一段,
     天頂仍然是藍的 —— 整片塗橘會變成濾鏡,不是天空。 */
  night:  {top:"#111d36", bot:"#2e3f63", mix:.06},
  dawn:   {top:"#4f7fb8", bot:"#f2a26e", mix:.10},
  day:    {top:"#5eb1ea", bot:"#d3e9f9", mix:.14},   /* 以前混了三成多的紙色,藍天變成灰白 */
  dusk:   {top:"#3f6396", bot:"#e5733d", mix:.08},
  evening:{top:"#25325a", bot:"#7a6a8e", mix:.08},
};
/* 天氣蓋過時段:暖色要壓掉,但天空本身還是藍的 —— 灰的是雲,
   不是天。純中性灰會變成黑白照片。 */
const WX_SKY={
  rain:  {top:"#5b7390", bot:"#9fb0c0", mix:.12},
  storm: {top:"#354357", bot:"#61707f", mix:.08},
  snow:  {top:"#7d93aa", bot:"#c3cfd9", mix:.16},
  cloud: {top:"#7fb0d6", bot:"#d3e0ea", mix:.14},   /* 多雲仍是藍天,只是淡一點;灰的是雲,不是天 */
};
function skyColour(part){
  const paper=pal("--panel","#fffdf7");
  const ph=skyPhase(localHour());
  const base=SKY[ph]||SKY.day;
  const wx=(typeof WX!=="undefined") && WX_SKY[WX.mode];
  if(!wx) return blendHex(paper, base[part], 1-base.mix);
  /* 有天氣時用天氣的灰,但夜裡仍然要更暗 —— 陰天的半夜不是灰白的。 */
  const dark=(ph==="night"||ph==="evening");
  const grey=blendHex(paper, wx[part], 1-wx.mix);
  return dark ? blendHex(grey, base[part], .55) : grey;
}

/* 天色暗的時候,深棕線條會整個消失在深藍裡。
   天上的東西(雲、鳥、山)改用淺色描,地面的維持深色。 */
function darkSky(){
  const p=skyPhase(localHour());
  /* 雷雨天再亮也是暗的,線條要用淺色才看得見 */
  return p==="night"||p==="evening"
      || (typeof WX!=="undefined"&&(WX.mode==="storm"||WX.mode==="rain"));
}
function skyInk(){ return darkSky() ? "#e8e2d2" : INK; }

function wxDim(){
  return WX.mode==="storm"?.58 : WX.mode==="rain"?.72
       : WX.mode==="snow"?.86 : WX.mode==="cloud"?.9 : 1;
}

/* 太陽或月亮:依當地時間決定位置與樣子 */
function drawSunMoon(g,W,H){
  const h=localHour(), ph=skyPhase(h);
  const night = ph==="night"||ph==="evening";
  /* 白天 6-18 點走一條弧線,夜裡月亮走另一條 */
  const t = night ? ((h<6? h+6 : h-18)/12) : ((h-6)/12);
  const x = 40 + Math.max(0,Math.min(1,t))*(W-80);
  const y = 46 - Math.sin(Math.max(0,Math.min(1,t))*Math.PI)*26;
  g.save();
  if(night){
    g.strokeStyle="#e8e2d2"; g.lineWidth=1.6;
    skCircle(g,x,y,9,5,"#f4f0e0");
    g.fillStyle=skyColour("top");
    g.beginPath(); g.arc(x+4.5,y-3,8.5,0,7); g.fill();
    /* 幾顆星,位置固定不閃 */
    g.fillStyle="#fffdf7";
    for(let i=0;i<9;i++){
      const sx=(W*0.08)+((i*137)%Math.max(1,W-60)), sy=18+((i*53)%40);
      g.globalAlpha=.5+((i*7)%5)/10;
      g.beginPath(); g.arc(sx,sy,1.2,0,7); g.fill();
    }
    g.globalAlpha=1;
  }else{
    const warm = (ph==="dawn"||ph==="dusk") ? "#e8683c" : "#f5c451";
    const tt=(typeof SCENE!=="undefined"&&SCENE.t)||0;
    /* 光暈,再加一圈一長一短慢慢呼吸的光芒 —— 靜止的太陽像貼紙。 */
    const glow=g.createRadialGradient(x,y,9,x,y,36);
    glow.addColorStop(0,"rgba(255,214,102,.55)"); glow.addColorStop(1,"rgba(255,214,102,0)");
    g.fillStyle=glow; g.beginPath(); g.arc(x,y,36,0,7); g.fill();
    g.strokeStyle="#e8a33c"; g.lineWidth=2.2; g.lineCap="round";
    for(let k=0;k<8;k++){
      const a=k*Math.PI/4-0.3+Math.sin(tt/60)*0.05, len=(k%2?20:24)+Math.sin(tt/22+k)*1.6;
      g.beginPath();
      g.moveTo(x+Math.cos(a)*15,y+Math.sin(a)*15);
      g.lineTo(x+Math.cos(a)*len,y+Math.sin(a)*len); g.stroke();
    }
    g.strokeStyle=INK; g.lineWidth=1.8;
    skCircle(g,x,y,11,5,warm);
  }
  g.restore();
}

/* 打雷:每隔一段時間閃一次。FLASH 記著這一道雷是第幾幀開始的,
   場景和貓都讀它 —— 貓要在同一瞬間被嚇到,不是各閃各的。 */
let FLASH=-999, NEXT_BOLT=0, BOLT_X=0;
function stormTick(t){
  if(WX.mode!=="storm"){ FLASH=-999; NEXT_BOLT=t+120; return; }
  if(t>=NEXT_BOLT){
    FLASH=t;
    BOLT_X=0.15+Math.random()*0.7;
    NEXT_BOLT=t+180+Math.floor(Math.random()*260);
  }
}
/* 這一瞬間正在閃嗎 */
function flashOn(t){ const d=t-FLASH; return d>=0 && d<10; }

/* 雨、雪、烏雲 */
function drawWeather(g,W,H,t){
  if(WX.mode==="clear") return;
  const FLR=H-22;
  if(WX.mode==="cloud"||WX.mode==="rain"||WX.mode==="snow"||WX.mode==="storm"){
    /* 雲是實心的、有底影、會飄。半透明的雲會把裡面的線透出來,看起來像一團亂線。
       多雲是淺灰白,雨雪雷是深灰。 */
    g.save();
    const white="#fbfaf4", soft=(WX.mode==="cloud");   /* 不用主題紙色:深色主題會把雲畫成黑的 */
    g.fillStyle=soft?blendHex(white,"#aeb9c6",.5):blendHex(white,"#76818f",.62);
    g.strokeStyle=skyInk(); g.lineWidth=1.5;
    const set=soft?[[0.08,.9,24],[0.24,.7,16],[0.40,1,30],[0.58,.75,20],[0.74,1.05,28],[0.90,.8,22]]
                  :[[0.12,1,26],[0.30,.85,18],[0.47,1.1,30],[0.66,.9,20],[0.84,1.05,26]];
    for(const [cp,cs,cy] of set){
      const cx=((W*cp + t*0.07*cs) % (W+120)) - 60;
      cloudPath(g,[[cx,cy,24*cs,9*cs],[cx-15*cs,cy+4*cs,14*cs,7*cs],[cx+16*cs,cy+4*cs,13*cs,6*cs]]);
      g.save(); g.globalAlpha=.35; g.fillStyle=soft?"#8a97a8":"#404b58";
      g.beginPath(); g.ellipse(cx,cy+6.5*cs,17*cs,2.6*cs,0,0,7); g.fill(); g.restore();
    }
    g.restore();
  }
  if(WX.mode==="rain"||WX.mode==="storm"){
    g.save(); g.strokeStyle="#7fa8c9"; g.lineWidth=1.4; g.lineCap="round";
    for(let i=0;i<70;i++){
      const sx=(i*97)%W, off=((t*3)+(i*31))%(FLR-40);
      const sy=40+off;
      g.globalAlpha=.55;
      g.beginPath(); g.moveTo(sx,sy); g.lineTo(sx-2,sy+9); g.stroke();
    }
    g.restore();
  }
  if(WX.mode==="storm"){
    const d=t-FLASH;
    if(d>=0 && d<10){
      /* 整片天空白一下,越後面越淡 */
      g.save();
      g.globalAlpha=(1-d/10)*0.55;
      g.fillStyle="#fffdf7";
      g.fillRect(4,4,W-8,H-8);
      g.restore();
    }
    if(d>=0 && d<7){
      /* 閃電本體:一條折線,不是直線 */
      const bx=W*BOLT_X;
      g.save();
      g.globalAlpha=1-d/7;
      g.strokeStyle="#fff4c2"; g.lineWidth=3.4;
      g.lineJoin="round"; g.lineCap="round";
      g.beginPath();
      g.moveTo(bx,18);
      g.lineTo(bx-9,44); g.lineTo(bx+5,48);
      g.lineTo(bx-7,78); g.lineTo(bx+8,74);
      g.lineTo(bx-3,104);
      g.stroke();
      g.strokeStyle="#e8a33c"; g.lineWidth=1.3; g.stroke();
      g.restore();
    }
  }
  if(WX.mode==="snow"){
    g.save(); g.fillStyle="#fffdf7";
    for(let i=0;i<55;i++){
      const drift=Math.sin((t/28)+i)*7;
      const sx=((i*113)%W)+drift, off=((t*0.9)+(i*37))%(FLR-40);
      g.globalAlpha=.75;
      g.beginPath(); g.arc(sx,40+off,1.9,0,7); g.fill();
    }
    g.restore();
    /* 地上積雪 */
    g.save(); g.globalAlpha=.8; g.fillStyle="#fffdf7";
    g.fillRect(10,FLR-3,W-20,3); g.restore();
  }
}

function blendHex(a,b,t){
  const p=c=>[parseInt(c.slice(1,3),16),parseInt(c.slice(3,5),16),parseInt(c.slice(5,7),16)];
  const [r1,g1,b1]=p(a),[r2,g2,b2]=p(b);
  const f=(x,y)=>Math.round(x+(y-x)*t).toString(16).padStart(2,"0");
  return "#"+f(r1,r2)+f(g1,g2)+f(b1,b2);
}

/* 雲的畫法:三團橢圓先描粗線、再整個填白 —— 填色蓋掉裡面那一半線,剩下的就是
   外圍一圈輪廓。以前是三個 ellipse 接在同一條路徑上直接 fill+stroke,ellipse 會從
   上一團的終點拉一條直線到下一團的起點,三團之間全是線,雲看起來像一團亂線。 */
function cloudPath(g,puffs){
  const path=()=>{ g.beginPath(); for(const [x,y,rx,ry] of puffs){ g.moveTo(x+rx,y); g.ellipse(x,y,rx,ry,0,0,7); } };
  const lw=g.lineWidth; g.lineWidth=lw*2; path(); g.stroke(); g.lineWidth=lw;
  path(); g.fill();
}
/* 場景用的五個顏色:旗子、小花、熱氣球輪流用。跟主題色不衝,但看得出是彩色的。 */
const YARD_HUES=["#e0645a","#f5c451","#4f8fcf","#5fae7a","#b57edc"];

/* Tally:寬扁的戶外貨場。左邊是招牌寫 TALLY 的棚子,中間貨櫃堆,
   右邊留空給走動的角色,太陽在右上。地平線壓在很低的位置。
   繽紛版:雲會飄、鳥成群飛、山是綠草丘長滿小花、棚頂掛彩旗、貨櫃紅藍綠、
   天上有熱氣球、地上有蝴蝶。(貨車試過,主人說不要。)
   全部照 t 推,沒有另外的狀態;夜裡和下雨把熱氣球、蝴蝶、鳥收起來。 */
function drawTallyYard(g,W,H,t){
  t=t||0;
  const FLR=H-14;
  const ink=INK, paper=pal("--panel","#fffdf7"), line=pal("--line","#caa356");
  const ph=skyPhase(localHour()), night=(ph==="night"||ph==="evening");
  const wet=(typeof WX!=="undefined")&&(WX.mode==="rain"||WX.mode==="storm"||WX.mode==="snow");

  drawSunMoon(g,W,H);

  /* 熱氣球:慢慢從左往右飄,輕輕上下浮。只在棚子右邊的天空,夜裡和壞天氣不飛。 */
  if(!night&&!wet){
    const span=W*0.6+140, bxp=W*0.4-70+((t*0.09)%span), byp=40+Math.sin(t/70)*4, R=13;
    if(bxp>W*0.40+8){
      g.save(); g.strokeStyle=ink; g.lineWidth=1.4;
      g.beginPath(); g.arc(bxp,byp,R,0,7); g.closePath();
      g.save(); g.clip();
      for(let k=0;k<5;k++){ g.fillStyle=YARD_HUES[k]; g.fillRect(bxp-R+k*(2*R/5),byp-R,2*R/5+0.6,2*R); }
      g.restore();
      g.beginPath(); g.arc(bxp,byp,R,0,7); g.stroke();
      skLine(g,bxp-5,byp+R-2,bxp-3,byp+R+9,301); skLine(g,bxp+5,byp+R-2,bxp+3,byp+R+9,302);
      skRect(g,bxp-5,byp+R+9,10,6,303,blendHex(paper,line,.6));
      g.restore();
    }
  }

  /* 雲:白的、底下一道淡藍陰影,慢慢往右飄,大的飄得快一點(近的動得多)。 */
  const sInk=skyInk();
  /* 雲是白的,跟主題的紙色無關 —— 深色主題的紙是深灰,拿來填雲就變成一團黑。 */
  const white="#fbfaf4";
  const cloudFill = darkSky() ? blendHex(white,"#4a5a7a",.3) : white;
  const cloudShade= darkSky() ? blendHex(cloudFill,"#2a3a5a",.35) : blendHex(white,"#9ec3e6",.5);
  g.strokeStyle=sInk; g.lineWidth=1.5;
  for(const [cp,cy,cs] of [[0.42,30,.85],[0.55,20,.6],[0.68,34,.75],[0.80,22,.55],[0.90,32,.68]]){
    const cx=((W*cp + t*0.10*cs) % (W+90)) - 45;
    g.fillStyle=cloudFill; g.strokeStyle=sInk;
    cloudPath(g,[[cx-12*cs,cy+3*cs,11*cs,5.5*cs],[cx+13*cs,cy+3*cs,10*cs,5*cs],[cx,cy,19*cs,7*cs]]);
    g.save(); g.globalAlpha=.55; g.fillStyle=cloudShade;
    g.beginPath(); g.ellipse(cx,cy+5.5*cs,13*cs,2.2*cs,0,0,7); g.fill(); g.restore();
  }
  /* 鳥:一小群排成 V,慢慢往右飛,翅膀一拍一拍;夜裡不飛。 */
  if(!night){
    g.strokeStyle=sInk; g.lineWidth=1.4; g.lineCap="round";
    const lead=((W*0.60 + t*0.32) % (W+80)) - 40, ly=50+Math.sin(t/45)*3;
    for(const [i,[ox,oy]] of [[0,0],[-15,7],[15,7],[-29,13],[29,13]].entries()){
      const bx=lead+ox, by=ly+oy, fl=Math.sin(t/7+i*0.7)*5;
      g.beginPath(); g.moveTo(bx-6,by); g.quadraticCurveTo(bx,by-fl,bx+6,by); g.stroke();
    }
  }
  /* 地面:地平線以下鋪一層淺草綠,不再是天空的顏色一路到底。 */
  g.fillStyle = night ? blendHex(paper,"#3a5a48",.45) : blendHex(paper,"#a9c77f",.55);
  g.fillRect(4,FLR,W-8,H-4-FLR);
  /* 綠草丘壓在地平線上、很扁,畫在角色之前 —— 高過人物就會像穿過去。丘上長小花,隨風輕搖。 */
  const grass = night ? blendHex(paper,"#3f6b4f",.55) : blendHex(paper,"#7cb36b",.6);
  const hills=[[0.58,180,15],[0.76,130,11],[0.92,150,13]];
  for(const [hp,hw,hh] of hills){
    const hx=W*hp;
    g.fillStyle=grass;
    g.beginPath(); g.moveTo(hx-hw/2,FLR);
    g.quadraticCurveTo(hx,FLR-hh*2.2,hx+hw/2,FLR);
    g.closePath(); g.fill(); g.strokeStyle=ink; g.lineWidth=1.3; g.stroke();
  }
  g.lineWidth=1.1;
  hills.forEach(([hp,hw,hh],hi)=>{
    const hx=W*hp;
    for(let k=0;k<6;k++){
      const u=(k+0.5)/6, fx=hx-hw/2+u*hw, fy=FLR-4.4*hh*u*(1-u);   /* 丘面(二次曲線)上的點 */
      const sway=Math.sin(t/38+k*1.3+hi)*1.2;
      g.strokeStyle=blendHex(grass,"#2f6b3d",.5);
      g.beginPath(); g.moveTo(fx,fy+1); g.lineTo(fx+sway,fy-7); g.stroke();
      g.fillStyle=YARD_HUES[(k+hi)%YARD_HUES.length];
      g.beginPath(); g.arc(fx+sway,fy-9,3.1,0,7); g.fill();
      g.fillStyle="#fff7d6"; g.beginPath(); g.arc(fx+sway,fy-9,1.1,0,7); g.fill();
    }
  });

  /* 棚子:左邊三分之一,寬扁,屋頂略拱 */
  /* 棚子要明顯高過人:工人連安全帽約 65px,所以棚身給 100。 */
  const bx0=14, bw=Math.min(400,W*0.36), bh=100, by=FLR-bh;
  g.strokeStyle=ink; g.lineWidth=2;
  skRect(g,bx0,by,bw,bh,31,blendHex(paper,line,.25));
  g.fillStyle=blendHex(paper,line,.45);
  g.beginPath();
  g.moveTo(bx0-10,by+5);
  g.quadraticCurveTo(bx0+bw/2,by-22,bx0+bw+10,by+5);
  g.lineTo(bx0+bw+10,by+11);
  g.quadraticCurveTo(bx0+bw/2,by-14,bx0-10,by+11);
  g.closePath(); g.fill(); g.stroke();
  /* 屋簷下一道淡影,牆面才有厚度 */
  g.save(); g.globalAlpha=.10; g.fillStyle="#000"; g.fillRect(bx0+1,by+11,bw-2,9); g.restore();
  /* 彩旗:兩根旗杆立在棚子兩端,一條繩子拉過屋頂,三角旗五色輪流、隨風輕擺。 */
  {
    const x1=bx0-6, x2=bx0+bw+6, top=by-24;
    g.lineWidth=1.6; skLine(g,x1,by+6,x1,top,311); skLine(g,x2,by+6,x2,top,312);
    g.lineWidth=1.2;
    g.beginPath(); g.moveTo(x1,top); g.quadraticCurveTo((x1+x2)/2,top+18,x2,top); g.stroke();
    const n=Math.max(6,Math.floor((x2-x1)/28));
    for(let k=0;k<n;k++){
      const u=(k+0.5)/n, px=x1+(x2-x1)*u, py=top+36*u*(1-u);
      const sway=Math.sin(t/26+k*0.9)*1.8;
      g.fillStyle=YARD_HUES[k%YARD_HUES.length];
      g.beginPath(); g.moveTo(px-6,py); g.lineTo(px+6,py); g.lineTo(px+sway,py+12); g.closePath(); g.fill(); g.stroke();
    }
    g.lineWidth=2;
  }
  /* 招牌 */
  const sgw=bw*0.52, sgx=bx0+bw*0.24;
  skRect(g,sgx,by+22,sgw,21,32,ink);
  g.fillStyle=pal("--tab","#ffb500"); g.textAlign="center";
  g.font='700 15px "Reprice Sketch","Iansui","Segoe Print","Ink Free",system-ui';
  g.fillText(SC.sign,sgx+sgw/2,by+38);
  g.textAlign="left";
  /* 捲門兩扇 */
  for(const d of [0.06,0.52]){
    const dx=bx0+bw*d, dw=bw*0.36, dy=by+52, dh=bh-52;
    skRect(g,dx,dy,dw,dh,33,paper);
    g.lineWidth=1.1;
    for(let k=1;k<4;k++) skLine(g,dx,dy+k*dh/4,dx+dw,dy+k*dh/4,33+k);
    g.lineWidth=2;
  }
  /* 兩扇門中間一盞壁燈:白天是一顆暖黃的燈,夜裡亮起來、有一圈光。 */
  {
    const lx=bx0+bw*0.47, ly=by+62;
    if(night){ const lg=g.createRadialGradient(lx,ly,2,lx,ly,26); lg.addColorStop(0,"rgba(255,214,102,.75)"); lg.addColorStop(1,"rgba(255,214,102,0)");
      g.fillStyle=lg; g.beginPath(); g.arc(lx,ly,26,0,7); g.fill(); }
    g.lineWidth=1.4; skLine(g,lx,ly-9,lx,ly-4,71);
    g.fillStyle=ink; g.fillRect(lx-5,ly-11,10,3);
    skCircle(g,lx,ly,4.2,72,night?"#ffd666":"#f5c451");
    g.lineWidth=2;
  }
  /* 棚子兩側各一叢矮灌木 */
  for(const [sx,sd] of [[bx0-2,73],[bx0+bw+14,74]]){
    g.lineWidth=1.3; g.fillStyle=night?blendHex(paper,"#2f5a3f",.6):blendHex(paper,"#5fae5a",.65);
    g.beginPath(); g.moveTo(sx-11,FLR); g.arc(sx-6,FLR-6,6,Math.PI,0); g.arc(sx+2,FLR-8,7,Math.PI,0); g.arc(sx+9,FLR-5,5,Math.PI,0); g.lineTo(sx+14,FLR); g.closePath(); g.fill(); g.strokeStyle=ink; g.stroke();
    g.lineWidth=2;
  }

  /* 貨櫃堆:兩疊,第一疊下藍上紅、第二疊綠 —— 貨場本來就是彩色的。一個貨櫃跟人差不多高。 */
  const CONT=[["#4f8fcf","#e0645a"],["#5fae7a"]];
  g.lineWidth=1.9;
  [[bx0+bw+34,2],[bx0+bw+210,1]].forEach(([cx,rows],si)=>{
    for(let r=0;r<rows;r++){
      const cw=160, ch=52, cy2=FLR-ch*(r+1)-2, col=CONT[si][r];
      skRect(g,cx,cy2,cw,ch,40+r,col);
      /* 浪板直紋 */
      g.lineWidth=1.1;
      for(let k=1;k<9;k++) skLine(g,cx+k*cw/9,cy2+5,cx+k*cw/9,cy2+ch-5,45+k+r*9);
      /* 上下橫樑,少了這個會像一塊木板 */
      g.lineWidth=1.6;
      skLine(g,cx+2,cy2+6,cx+cw-2,cy2+6,60+r);
      skLine(g,cx+2,cy2+ch-6,cx+cw-2,cy2+ch-6,62+r);
      /* 門把 */
      g.lineWidth=1.4;
      skLine(g,cx+cw*0.5,cy2+12,cx+cw*0.5,cy2+ch-12,64+r);
      /* 白色的小名牌、四角的角件、頂邊一道亮光、最底下一道影子 —— 貨櫃才像鐵的 */
      g.lineWidth=1.1; skRect(g,cx+cw*0.62,cy2+ch*0.34,30,11,66+r,paper);
      g.fillStyle=blendHex(col,"#000000",.45);
      for(const [qx,qy] of [[cx+2,cy2+2],[cx+cw-8,cy2+2],[cx+2,cy2+ch-8],[cx+cw-8,cy2+ch-8]]) g.fillRect(qx,qy,6,6);
      g.fillStyle=blendHex(col,"#ffffff",.35); g.fillRect(cx+9,cy2+1.5,cw-18,2.5);
      if(r===0){ g.save(); g.globalAlpha=.16; g.fillStyle="#000"; g.fillRect(cx-3,FLR-3,cw+6,3); g.restore(); }
      g.lineWidth=1.9;
    }
  });
  /* 地平線 */
  g.strokeStyle=ink; g.lineWidth=1.8;
  skLine(g,10,FLR,W-10,FLR,3);

}
/* ---------------------------------------------------------------------
   蝴蝶
   ---------------------------------------------------------------------
   兩隻,各有自己的狀態:fly 在草丘前繞 → to 挑一隻貓或狗飛過去 → sit 停在頭頂
   → off 飛走。停著的時候動物停下來不走:狗抬頭看、耳朵豎起、尾巴搖得飛快,最後
   甩甩頭把它甩走;貓抖耳朵、眼睛往上盯,最後瞪大眼睛看著它飛走。手不動。
   動物身上只掛兩個欄位:bug(哪一隻停著/正飛過來)和 bugAt(幾時停的)。
   --------------------------------------------------------------------- */
function bugsInit(){
  const W=SCENE.W||900, H=SCENE.H||190;
  SCENE.bugs=[0,1].map(i=>({i, x:W*(0.60+i*0.13), y:H-60, act:"fly", wait:240+i*300, tgt:null, vx:0,
    c1:i?"#4aa3df":"#ff8c42", c2:i?"#f5c451":"#e85d75"}));
}
/* 狗興奮時的姿勢:蝴蝶飛近就原地小跳;停久了甩頭。畫狗和算頭頂位置都用這一份。 */
function dogPose(a,t){
  const b=a.bug;
  const chase=b&&b.act==="to"&&Math.hypot(b.x-a.x,b.y-(a.y-40))<70;
  const sit=b&&b.act==="sit";
  return {hop: chase?Math.abs(Math.sin(t/6))*5:0, shake: (sit&&b.wait<40)?Math.sin(t/1.5)*3:0, excited:!!(chase||sit), sit:!!sit};
}
/* 頭頂:蝴蝶停的位置。跟 drawDog / drawCat 算頭的方式一致。 */
function bugHead(a,t){
  const f=a.dir<0?-1:1;
  if(a.kind==="dog"){ const p=dogPose(a,t); return {x:a.x+17*f+p.shake, y:a.y-p.hop-36-13-4}; }
  const walking=a.act==="walk", hxo=walking?14*f:0, hy=walking?a.y-24:a.y-(a.act==="puff"?38:34);
  return {x:a.x+hxo, y:hy-11-4};
}
function bugsStep(){
  const B=SCENE.bugs; if(!B||BRAND!=="Tally") return;
  const W=SCENE.W||900, FLR=(SCENE.H||190)-14, t=SCENE.t;
  B.forEach((b,i)=>{
    if(b.act==="fly"){
      const cx=W*(0.60+i*0.13)+Math.sin(t/95+i*2)*70, cy=FLR-46+Math.sin(t/21+i)*9;
      b.x+=(cx-b.x)*0.04; b.y+=(cy-b.y)*0.04;
      if(--b.wait<=0){
        const pets=SCENE.actors.filter(a=>(a.kind==="dog"||a.kind==="cat")&&!a.bug&&a.act!=="pee"&&a.act!=="puff"&&a.act!=="bark");
        if(pets.length){ b.tgt=pets[(i+Math.floor(t/7))%pets.length]; b.tgt.bug=b; b.act="to"; b.wait=900; }
        else b.wait=120;
      }
    }else if(b.act==="to"){
      const h=bugHead(b.tgt,t), dx=h.x-b.x, dy=h.y-b.y, d=Math.hypot(dx,dy);
      if(d<2.5){
        b.act="sit"; b.wait=260+((i*131)%160); b.tgt.bugAt=t;
        /* 動物停下來,而且要等蝴蝶走了才動 */
        b.tgt.act=b.tgt.kind==="cat"?"sit":"idle"; b.tgt.wait=Math.ceil(b.wait/3)+20;
      }else{
        b.x+=dx/d*1.3+Math.sin(t/4)*0.6; b.y+=dy/d*1.3+Math.cos(t/5)*0.5;
        if(--b.wait<=0){ b.tgt.bug=null; b.tgt=null; b.act="fly"; b.wait=300; }
      }
    }else if(b.act==="sit"){
      const h=bugHead(b.tgt,t); b.x=h.x; b.y=h.y;
      const a=b.tgt, scared=a.act==="walk"||a.act==="puff"||a.act==="pee"||a.act==="bark";
      /* 時間到就飛走(狗甩頭甩到底、貓瞪完);動物自己走掉或嚇到也飛。 */
      if(--b.wait<=0 || scared){ b.act="off"; b.wait=70; b.vx=(i?-1:1)*0.9; }
    }else{
      b.x+=b.vx+Math.sin(t/4)*0.8; b.y-=1.4;
      if(--b.wait<=0){ if(b.tgt){ b.tgt.bug=null; b.tgt.bugAt=0; } b.tgt=null; b.act="fly"; b.wait=420+((i*97)%300); }
    }
  });
}
function drawBug(g,b,t){
  const sit=b.act==="sit";
  const flap=sit?0.55+0.35*Math.sin(t/16):Math.abs(Math.sin(t/5+b.i))*0.7+0.3;
  g.save(); g.translate(b.x,b.y); g.scale(1.7,1.7); g.strokeStyle=INK; g.lineWidth=0.8;
  for(const sd of [-1,1]){
    g.fillStyle=b.c1; g.beginPath(); g.ellipse(sd*4*flap,-2,4*flap,3.2,0,0,7); g.fill(); g.stroke();
    g.fillStyle=b.c2; g.beginPath(); g.ellipse(sd*3*flap,2.5,3*flap,2.2,0,0,7); g.fill(); g.stroke();
  }
  g.fillStyle=INK; g.fillRect(-0.6,-4,1.2,8);
  g.lineWidth=0.6; g.beginPath(); g.moveTo(0,-4); g.lineTo(-2.2,-7.5); g.moveTo(0,-4); g.lineTo(2.2,-7.5); g.stroke();
  g.restore();
}
function sceneDraw(){
  const g=SCENE.ctx; if(!g) return;
  const W=SCENE.W,H=SCENE.H,t=SCENE.t;
  g.clearRect(0,0,W,H);
  g.lineWidth=2.2; g.strokeStyle=INK;

  if(BRAND==="Tally"){
    /* 天空要漸層才畫得出「天頂藍、地平線橘」。單色只能二選一。 */
    const sky=g.createLinearGradient(0,3,0,H-22);
    sky.addColorStop(0, skyColour("top"));
    sky.addColorStop(0.62, blendHex(skyColour("top"),skyColour("bot"),.55));
    sky.addColorStop(1, skyColour("bot"));
    skRect(g,3,3,W-6,H-6,101,sky);
    drawTallyYard(g,W,H,t);
    drawWeather(g,W,H,t);
  }else{
    /* Scoobi:室內倉庫。招牌、窗、左邊貨架、橫貫全寬的輸送帶、右邊堆高機。 */
    skRect(g,3,3,W-6,H-6,101,pal("--bg","#f4ecdc"));
    const FLR=H-14, paper=pal("--panel","#fffdf7");
    g.strokeStyle=INK; g.lineWidth=2;
    skRect(g,26,14,140,30,20,paper);
    g.fillStyle=INK; g.textAlign="center";
    g.font='700 16px "Reprice Sketch","Iansui","Segoe Print","Ink Free",system-ui';
    g.fillText("Scoobi",96,35); g.textAlign="left";
    g.lineWidth=1.6;
    for(let k=0;k<4;k++){
      const wx=200+k*64;
      const wsky=g.createLinearGradient(0,16,0,44);
      wsky.addColorStop(0, skyColour("top"));
      wsky.addColorStop(1, skyColour("bot"));
      skRect(g,wx,16,40,28,30+k,wsky);
      skLine(g,wx+20,16,wx+20,44,34+k);
      skLine(g,wx,30,wx+40,30,38+k);
    }
    const shx=20, shw=104, shy=FLR-78;
    g.lineWidth=1.8;
    for(let r=0;r<3;r++){
      const ry=shy+r*26;
      skLine(g,shx,ry+24,shx+shw,ry+24,52+r);
      for(let b=0;b<3;b++) skRect(g,shx+4+b*33,ry+4,28,19,55+r*3+b,BOX);
    }
    skLine(g,shx,shy,shx,FLR,60); skLine(g,shx+shw,shy,shx+shw,FLR,61);
    /* 輸送帶要高過角色頭頂,不然帶子會從貓狗身上穿過去。
       貓最高約 45px(身體 30 + 頭 15),所以帶底至少留 52。 */
    const by=FLR-72, bx0=shx+shw+22, bx1=W-120;
    g.lineWidth=2;
    skRect(g,bx0,by,bx1-bx0,11,70,blendHex(INK,paper,.3));
    if(SCENE.progress>0){
      g.fillStyle=pal("--tab","#ffb500");
      g.fillRect(bx0+2,by+2,(bx1-bx0-4)*Math.min(1,SCENE.progress),7);
    }
    g.fillStyle=pal("--tab","#ffb500");
    for(let x=bx0+30+((t/3)%36); x<bx1-8; x+=36){
      g.beginPath(); g.arc(x,by+5.5,2.2,0,7); g.fill();
    }
    g.strokeStyle=INK; g.lineWidth=2;
    for(const lx of [bx0+30,(bx0+bx1)/2,bx1-30]) skLine(g,lx,by+11,lx,FLR,72);
    g.lineWidth=1.8;
    skRect(g,W-66,FLR-64,42,64,80,paper);
    for(let r=1;r<4;r++) skLine(g,W-66,FLR-64+r*16,W-24,FLR-64+r*16,82+r);
    skRect(g,W-108,FLR-24,28,18,86,pal("--tab","#ffb500"));
    skRect(g,W-112,FLR-40,18,16,87,BOX);
    g.fillStyle=INK;
    for(const wx of [W-102,W-86]){ g.beginPath(); g.arc(wx,FLR,4,0,7); g.fill(); }
    g.strokeStyle=INK; g.lineWidth=1.8;
    skLine(g,10,FLR,W-10,FLR,3);
    /* 窗外的天氣:只畫在窗格範圍內 */
    g.save();
    g.beginPath();
    for(let k=0;k<4;k++) g.rect(200+k*64,16,40,28);
    g.clip();
    drawSunMoon(g,W,60);
    drawWeather(g,W,80,t);
    g.restore();
    g.fillStyle=INK;
    g.font='600 13px "Reprice Sketch","Iansui","Segoe Print","Ink Free",system-ui';
    g.fillText((SC.lang==="en")?"Tap the cat or the dog!":"\u9ede\u4e00\u4e0b\u8c93\u6216\u72d7\uff01", bx0+8, FLR-64);
  }

  for(const a of SCENE.actors){
    if(a.kind==="worker") drawWorker(g,a,t);
    else if(a.kind==="dog") drawDog(g,a,t);
    else drawCat(g,a,t);
  }
  /* 蝴蝶最後畫:停在頭上時要在頭的前面。 */
  if(BRAND==="Tally"&&SCENE.bugs) for(const b of SCENE.bugs) drawBug(g,b,t);
}
function sceneLoop(){
  /* stormTick 要在 sceneStep 之前:貓讀 FLASH 決定要不要嚇到,
     順序反過來的話貓永遠比閃電慢一幀。 */
  SCENE.t++; stormTick(SCENE.t); sceneStep(); bugsStep(); sceneDraw();
  SCENE.raf=requestAnimationFrame(sceneLoop);
}

  window.RMAScene = {
    init(opts) {
      Object.assign(SC, opts || {});
      sceneInit();
      sceneShow(true);
    },
    setLang(l) { SC.lang = l; if (SCENE.cv) SCENE.cv.title = SC.lang === "en" ? "Click an animal" : "點一下摸摸"; },
    setTz(z) { SC.tz = z; },
    setProgress(p) { SCENE.progress = Math.max(0, Math.min(1, +p || 0)); },
    setWeather(m) { WX.mode = m || "clear"; wxSave(); },
    show(on) { sceneShow(on); },
    resize() { if (SCENE.cv) { sceneResize(); spreadActors(); sceneDraw(); } },
  };
})();
