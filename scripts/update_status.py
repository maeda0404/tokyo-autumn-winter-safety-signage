import json,re,urllib.request
from datetime import datetime,timezone,timedelta
from pathlib import Path
JST=timezone(timedelta(hours=9)); OUT=Path(__file__).resolve().parents[1]/'data/status.json'
def get(url,json_mode=True):
 r=urllib.request.urlopen(urllib.request.Request(url,headers={'User-Agent':'safety-signage/1.0'}),timeout=30); b=r.read(); return json.loads(b) if json_mode else b.decode('utf-8','replace')
def main():
 old=json.loads(OUT.read_text(encoding='utf-8')); errors=[]
 try:
  w=get('https://www.jma.go.jp/bosai/warning/data/warning/130000.json'); text=json.dumps(w,ensure_ascii=False)
  wind=('強風' in text or '暴風' in text) and '解除' not in text; dry='乾燥' in text and '解除' not in text
 except Exception: wind=old['weather']['wind']['active']; dry=old['weather']['dry']['active']; errors.append('注意報')
 try:
  f=get('https://www.jma.go.jp/bosai/forecast/data/forecast/130000.json'); vals=[]
  for rep in f:
   for ts in rep.get('timeSeries',[]):
    for a in ts.get('areas',[]):
     if str(a.get('area',{}).get('code'))=='44132':
      for v in a.get('temps',[]):
       try: vals.append(float(v))
       except: pass
  low=min(vals) if vals else None; cold=low is not None and low<=3
 except Exception: low=None; cold=old['weather']['cold']['active']; errors.append('気温')
 try:
  h=get('https://idsc.tmiph.metro.tokyo.lg.jp/diseases/flu/flu/',False); t=re.sub('<[^>]+>',' ',h)
  def get_flu_level(value):
    if value >= 30:
        return "警報レベル"

    if value >= 10:
        return "注意報レベル"

    if value >= 1:
        return "流行中"

    return "非流行"
 except Exception: level='確認中'; errors.append('感染症')
 d={'updated':datetime.now(JST).strftime('%Y-%m-%d %H:%M'),'sourceStatus':'正常' if not errors else '一部取得失敗：'+'・'.join(errors),'influenza':{'level':level,'period':'東京都','trend':'公式発表を自動確認'},'weather':{'wind':{'active':wind,'normalText':'情報なし','alertText':'強風注意','note':'飛散・揚重確認'},'dry':{'active':dry,'normalText':'情報なし','alertText':'火気注意','note':'消火確認を徹底'},'cold':{'active':cold,'normalText':'情報なし','alertText':'凍結注意','note':f'予想最低気温 {low:g}℃' if low is not None else '通路・足元確認'}}}
 OUT.write_text(json.dumps(d,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
if __name__=='__main__': main()
