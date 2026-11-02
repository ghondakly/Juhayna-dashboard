"""build.py APP.html LEGACY.json|- MASTER.json|- OUT_ARTIFACT.html OUT_REPO.html"""
import sys
app,leg,mas,out_a,out_r=sys.argv[1:6]
s=open(app,encoding='utf-8').read()
L=open(leg,encoding='utf-8').read() if leg!='-' else 'null'
Mj=open(mas,encoding='utf-8').read() if mas!='-' else 'null'
s=s.replace('/*__BAKED__*/','const BAKED_LEGACY='+L+';const BAKED_MASTER_V2='+Mj+';',1)
head,body=s.split('<!--BODY-->',1)
open(out_a,'w',encoding='utf-8').write((head+body).replace('<title>Juhayna Fleet Control Tower</title>','<title>Control Tower v2 Preview</title>',1))
open(out_r,'w',encoding='utf-8').write('<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'+head+'</head>\n<body>\n'+body+'\n</body>\n</html>\n')
print("built",len(head+body)//1024,"KB")
