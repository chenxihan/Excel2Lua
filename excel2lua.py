#!/usr/bin/python
#-*- coding: UTF-8 -*-

import xlrd
import os
import math
import ConfigParser
import sys
import time
import re

isdebug = False
sleeptime = 600

reload(sys)
encoding="utf-8"
sys.setdefaultencoding(encoding)
config=ConfigParser.ConfigParser()
config.readfp(open("dir.ini"))
indir=config.get("path","IN")
serveroutdir=config.get("path","SERVER_OUT")
clientoutdir=config.get("path","CLIENT_OUT")
deep_set=int(config.get("path","DEEP"))
ALIAS_DICT={}
with open("./alias.txt") as alias:
	for line in alias.readlines():
		if not line.startswith("#"):
			line=line.rstrip('\n')
			if line!="":
				l=line.split("=")
				ALIAS_DICT[l[0]]=l[1].rstrip('\n')
#alias.txt
for x in ALIAS_DICT.keys():
	if isdebug:
		print(x,ALIAS_DICT[x])

def foreach_parse(val,sp,parser,attr):
	ret=[]
	if val=="":
		if attr.find("e")!=-1:
			return None
		return ret
	for v in str(val).split(sp):
		ret.append(parser(v,attr))
	return ret

def trans_unixtime(s):
	return int(time.mktime(time.strptime(s,"%Y%m%d%H%M%S")))

def try_parseint(s):
	try:
		return int(s)
	except ValueError as e:
		return int(round(float(s)))
		cnt+=1

def parse_effect_list(s):
	r={}
	d=-1
	i=0
	cnt=1
	while i<len(s):
		c=s[i]
		if c.isdigit():
			if d==-1:
				d=i
		elif c=="-":
			if d==-1:
				d=i
		elif c==",":
			if d!=-1:
				r[cnt]=try_parseint(s[d:i])
				cnt+=1
				d=-1
		elif c=="(":
			ni,o=parse_effect_list(s[i+1:])
			r[cnt]=o
			cnt+=1
			if d!=-1:
				o[0]=try_parseint(s[d:i])

			i=i+ni+1
			d=-1
		elif c==")":
			if d!=-1:
				r[cnt]=try_parseint(s[d:i])
				#cnt+=1
			#if cnt==2:
				#return i,r[1]
			return i,r
		i=i+1

	if d!=-1:
		r[cnt]=try_parseint(s[d:i])
		cnt+=1

	return i,r


def parse_effectcast(s,attr):
	i=0
	while i<len(s):
		c=s[i]
		if c.isdigit()!=True:
			break
		i+=1

	assert(s[i]=='(')
	assert(s[-1]==')')
	r={}
	if i!=0:
		c=try_parseint(s[:i])
		r[0]=c
	i,o=parse_effect_list(s[i+1:-1])
	r[1]=o
	return r


def parse_integer(s,attr):
	if not isinstance(s, float):
		if s=="" and attr.find("e")!=-1:
			return None
		if s[-1:]=='u':
			return trans_unixtime(s[0:-1])
		elif s[-1:]=='t':
			if attr.find('s')!=-1:
				s=s[:-1]
			else:
				s=s[:-1]+'0'
	if isdebug:
		print("def parse_integer : s=",s,"attr=",attr)
	#返回浮点数x的四舍五入值。
	return int(round(float(s)))

def parse_double(val,attr):
	if val=="" and attr.find("e")!=-1:
		return None
	if isdebug:
		print("def parse_double : val=",val,"attr=",attr)
	return float(val)

def parse_string(val,attr):
	if val=="" and attr.find("e")!=-1:
		return None
	if isdebug:
		print("def parse_string : val=",val,"attr=",attr)
	return str(val).encode(encoding)

def parser_any(val,attr):
	mobj=re.match(r"(.*?)\|(.*)",val)
	if mobj:
		parser=mobj.group(1)
		valn=mobj.group(2)
		if isdebug:
			print("def parser_any : val=",val,"attr=",attr)
		return get_parser(parser)(valn,attr)
	elif val=="" and attr.find("e")!=-1:
		if isdebug:
			print("def parser_any : return None,val=",val,"attr=",attr)
		return None
	else:
		print("Any type error")
		time.sleep(sleeptime)
		raise Exception("Any type error")


PARSER={}
PARSER["int"]=parse_integer
PARSER["double"]=parse_double
PARSER["string"]=parse_string
#PARSER["effectcast"]=parse_effectcast
#PARSER["any"]=parser_any
def get_parser(s):
	parser=PARSER.get(s,None)
	if parser:
		if isdebug:
			print("IN PARSER",s)
		return parser
	return build_array(s)

def build_array(s):
	alias_s=ALIAS_DICT.get(s,None)
	if alias_s:
		if isdebug:
			print("IN ALIAS_DICT:",s)
		s=alias_s
	mobj=re.match(r"array<(.*?),(.*)>",s)
	if mobj:
		sp=mobj.group(1)
		subcall=[]
		for v in str(mobj.group(2)).split(','):
			subcall.append(get_parser(v))
		def array_func(val,r1):
			r=[]
			if val=="":
				return r
			l=str(val).split(sp)
			for i in xrange(len(l)):
				if i<len(subcall):
					r.append(subcall[i](l[i],r1))
				else:
					r.append(subcall[-1](l[i],r1))
			return r
		return array_func
	else:
		print("Invalid type ",str(s))
		time.sleep(sleeptime)
		raise Exception("Invalid type "+str(s))


def septer(deep,lineding_deep):
	if deep>lineding_deep:
		return "",""
	else:
		return "\t","\n"

def thespace(n,sp):
	s=""
	for i in xrange(n):
		s=s+sp
	return s

def trans_list(obj,deep,lineding_deep):
	space,lending=septer(deep,lineding_deep)
	vals=[]
	#返回对象（字符、列表、元组等）长度或项目个数。
	objlen=len(obj)
	#xrange数组生成器
	for i in xrange(objlen):
		#v=thespace(deep,space)+"["+str(i+1)+"]="+trans_obj(obj[i],deep,lineding_deep)
		v=thespace(deep,space)+trans_obj(obj[i],deep,lineding_deep)
		vals.append(v)

	return "{"+lending+(","+lending).join(vals)+lending+thespace(deep-1,space)+"}"

numchar=("0","1","2","3","4","5","6","7","8","9")
def trans_dict(obj,deep,lineding_deep):
	space,lending=septer(deep,lineding_deep)
	keys=obj.keys()
	keys.sort()
	vals=[]
	for k in keys:
		val=obj[k]
		if val!=None:
			v=thespace(deep,space)
			if isinstance(k, int):
				v=v+"["+str(k)+"]="
			elif isinstance(k, str):
				try:
					float(k)
				except ValueError as e:
					if k.startswith(numchar) or k.find("%")!=-1:
						v=v+"['"+str(k)+"']="
					else:
						v=v+str(k)+"="
				else:
					v=v+"['"+str(k)+"']="
			else:
				print("Invalid key type ",k)
				time.sleep(sleeptime)
				raise Exception("Invalid key type")

			v=v+trans_obj(obj[k],deep,lineding_deep);
			vals.append(v)

	return "{"+lending+(","+lending).join(vals)+lending+thespace(deep-1,space)+"}"


def trans_obj(obj,deep,lineding_deep):

	space,lending=septer(deep,lineding_deep)
	# 函数来判断一个对象是否是int的类型，类似 type()。
	if isinstance(obj, int):
		#函数将对象转化为适于人阅读的形式,转string类型。
		if isdebug:
			print("int:",obj)
		return str(obj)
	elif isinstance(obj, long):
		if isdebug:
			print("long:",obj)
		return str(obj)
	elif isinstance(obj, float):
		if isdebug:
			print("float:",obj)
		return str(obj)
	elif isinstance(obj, str):
		#str.replace(old, new[, max])
		if isdebug:
			print("str:",obj)
		return "\""+obj.replace('\"','\\\"')+"\""
	elif isinstance(obj,list):
		if isdebug:
			print("list:",obj)
		return trans_list(obj,deep+1,lineding_deep)
	elif isinstance(obj,dict):
		if isdebug:
			print("dict:",obj)
		return trans_dict(obj,deep+1,lineding_deep)
	elif obj==None:
		if isdebug:
			print("None:",obj)
		return "nil"
	else:
		print("Error type:",str(type(obj)))
		time.sleep(sleeptime)
		raise Exception("Invalid obj "+str(type(obj)))

def trans2lua(out,obj,lineding_deep):
	out.write("return ")
	s=trans_obj(obj,0,lineding_deep)
	out.write(s)
	out.write("\n")

def transfer_z(sret,cret,booksheet,name):
	if isdebug:
		print("def transfer_z:",name,"row:",booksheet.nrows)
	if booksheet.nrows<4:
		print("Error format: filename:",name,booksheet.name.encode(encoding))
		time.sleep(sleeptime)
		raise Exception("Error format "+ name +"."+booksheet.name.encode(encoding))

	stb={}
	ctb={}
	cuse=False
	suse=False

	for row in xrange(booksheet.nrows):
		if(row<4):
			continue
		else:
			srt={}
			crt={}
			key=None
			for col in xrange(booksheet.ncols):
				coltype=booksheet.cell(1,col).value.encode(encoding)
				colname=booksheet.cell(2,col).value.encode(encoding)
				attr=booksheet.cell(3,col).value.encode(encoding)
				val=booksheet.cell(row,col).value
				if isdebug:
					print("row=",row+1,"col=",col+1,"coltype=",coltype,"colname=",colname,"attr=",attr,"val=",val,)
				rval=None
				if len(attr)==0:
					continue
				if attr.find("e")!=-1 and not val:
					continue

				try:
					parser=get_parser(coltype)
					if attr.find("k")!=-1:
						if key!=None:
							print("mult key using:",)
							time.sleep(sleeptime)
							raise Exception("mult key using!")
						rval=parser(val,attr.replace('c',''))
						key=rval
						crt[colname]=rval
						srt[colname]=rval

					if attr.find("c")!=-1:
						if attr.find("m")!=-1:
							rval=foreach_parse(val,"|",parser,attr.replace('s',''))
						else:
							rval=parser(val,attr.replace('s',''))
						crt[colname]=rval
						cuse=True

					if attr.find("s")!=-1:
						if attr.find("m")!=-1:
							rval=foreach_parse(val,"|",parser,attr.replace('c',''))
						else:
							rval=parser(val,attr.replace('c',''))
						srt[colname]=rval
						suse=True

				except Exception as e:
					import traceback
					print("Exception at "+name +"."+booksheet.name.encode(encoding)
						+"("+str(row+1)+","+str(col+1)+")\n"
						+repr(e)+"\n")
						#+traceback.format_exc()))
					time.sleep(sleeptime)
					raise Exception("Exception at "+name +"."+booksheet.name.encode(encoding)
						+"("+str(row)+","+str(col)+")\n"
						+repr(e)+"\n"
						+traceback.format_exc())
			if key==None:
				print("must set a key")
				time.sleep(sleeptime)
				raise Exception("must set a key")
			stb[key]=srt
			ctb[key]=crt

	if cuse:
		cret[booksheet.name.encode(encoding)]=ctb
	if suse:
		sret[booksheet.name.encode(encoding)]=stb

def transfer_y(sret,cret,booksheet,name):
	if booksheet.nrows<3:
		print("Error format: filename:",name,booksheet.name.encode(encoding))
		time.sleep(sleeptime)
		raise Exception("Error format "+ name +"."+booksheet.name.encode(encoding))

	key_coltype=booksheet.cell(1,0).value.encode(encoding)
	key_attr=booksheet.cell(2,0).value.encode(encoding)
	val_coltype=booksheet.cell(1,1).value.encode(encoding)
	val_attr=booksheet.cell(2,1).value.encode(encoding)
	val_mult=val_attr.find("m")!=-1
	cuse=key_attr.find("c")!=-1 or val_attr.find("c")!=-1
	suse=key_attr.find("s")!=-1 or val_attr.find("s")!=-1

	ctb={}
	stb={}
	for row in xrange(booksheet.nrows):
		if(row<3):
			continue
		else:
			val_k=booksheet.cell(row,0).value
			val_v=booksheet.cell(row,1).value

			parser_k=get_parser(key_coltype)
			parser_v=get_parser(val_coltype)

			if cuse:
				if val_mult:
					ctb[parser_k(val_k,key_attr.replace('s',''))]=foreach_parse(val_v,"|",parser_v,val_attr.replace('s',''))
				else:
					ctb[parser_k(val_k,key_attr.replace('s',''))]=parser_v(val_v,val_attr.replace('s',''))
			if suse:
				if val_mult:
					stb[parser_k(val_k,key_attr.replace('c',''))]=foreach_parse(val_v,"|",parser_v,val_attr.replace('c',''))
				else:
					stb[parser_k(val_k,key_attr.replace('c',''))]=parser_v(val_v,val_attr.replace('c',''))

	if cuse:
		cret[booksheet.name.encode(encoding)]=ctb
	if suse:
		sret[booksheet.name.encode(encoding)]=stb

def transfer(name):
	sret={}
	cret={}
	workbook=xlrd.open_workbook(name)
	
	for booksheet in workbook.sheets():
		booksheetname=booksheet.name.encode(encoding)
		if isdebug:
			print("booksheetname :",booksheetname)
		if booksheetname.startswith("y_"):
			transfer_y(sret,cret,booksheet,name)
			continue;

		if booksheetname.startswith("z_"):
			transfer_z(sret,cret,booksheet,name);
			continue;

	return sret,cret

def lua_test(files):
	for f in files:
		r=os.popen("lua "+f)

#获取路径indir下的所有文件
def grap_infs(indir):
	#要返回的对象。存储所有文件
	ret=[]
	#返回path指定的文件夹包含的文件或文件夹的名字的列表。
	fs=os.listdir(indir)
	for f in fs:
		fname=indir+"/"+f
		#如果是文件
		if os.path.isfile(fname):
			ret.append(fname)
		#如果是文件夹
		elif os.path.isdir(fname):
			ret.extend(grap_infs(fname))
	return ret

def main():
	#获取所有文件
	fs=grap_infs(indir)
	for fname in fs:
		if isdebug:
			print("===============>> file_name:"+fname)
		#生成的s、c的lua文件的路径
		luafs=[]
		#获取单个文件中的数据分别存在s和c的字典中
		s,c=transfer(fname)

		for k in s.keys():
			#创建文件的名字
			fn=serveroutdir+"/"+k[2:]+".lua"
			luafs.append(fn)
			if isdebug:
				print("========== s   filename: ",fn,"s key:",k)#,"s value:",s[k])
			#打开一个文件只用于写入。如果该文件已存在则打开文件，并从开头开始编辑，即原有内容会被删除。如果该文件不存在，创建新文件。
			out=open(fn,"w")
			#转成lua格式并写入
			trans2lua(out,s[k],deep_set)
			#关闭文件
			out.close()

		for k in c.keys():
			fn=clientoutdir+"/"+k[2:]+".lua"
			luafs.append(fn)
			if isdebug:
				print("========== c   filename: ",fn,"c key:",k)#,"c value:",c[k])

			out=open(fn,"w")
			trans2lua(out,c[k],deep_set)
			out.close()
		#用cmd的lua命令打开lua文件，测试其正确性
		lua_test(luafs)

		print("end......................")
		#time.sleep(sleeptime)
	return 0

if __name__=="__main__":
	main()