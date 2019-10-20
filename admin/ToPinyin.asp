<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Session.CodePage=936
Response.ContentType = "text/html"
Response.Charset = "GB2312"
Response.Expires=-999
Session.Timeout=999
'==========================================
'文 件 名：ToPinyin.asp
'文件用途：转换为拼音
'==========================================
Set Pinyin=CreateObject("Scripting.Dictionary") 
Pinyin.add "A",-20319 
Pinyin.add "Ai",-20317 
Pinyin.add "An",-20304 
Pinyin.add "Ang",-20295 
Pinyin.add "Ao",-20292 
Pinyin.add "Ba",-20283 
Pinyin.add "Bai",-20265 
Pinyin.add "Ban",-20257 
Pinyin.add "Bang",-20242 
Pinyin.add "Bao",-20230 
Pinyin.add "Bei",-20051 
Pinyin.add "Ben",-20036 
Pinyin.add "Beng",-20032 
Pinyin.add "Bi",-20026 
Pinyin.add "Bian",-20002 
Pinyin.add "Biao",-19990 
Pinyin.add "Bie",-19986 
Pinyin.add "Bin",-19982 
Pinyin.add "Bing",-19976 
Pinyin.add "Bo",-19805 
Pinyin.add "Bu",-19784 
Pinyin.add "Ca",-19775 
Pinyin.add "Cai",-19774 
Pinyin.add "Can",-19763 
Pinyin.add "Cang",-19756 
Pinyin.add "Cao",-19751 
Pinyin.add "Ce",-19746 
Pinyin.add "Ceng",-19741 
Pinyin.add "Cha",-19739 
Pinyin.add "Chai",-19728 
Pinyin.add "Chan",-19725 
Pinyin.add "Chang",-19715 
Pinyin.add "Chao",-19540 
Pinyin.add "Che",-19531 
Pinyin.add "Chen",-19525 
Pinyin.add "Cheng",-19515 
Pinyin.add "Chi",-19500 
Pinyin.add "Chong",-19484 
Pinyin.add "Chou",-19479 
Pinyin.add "Chu",-19467 
Pinyin.add "Chuai",-19289 
Pinyin.add "Chuan",-19288 
Pinyin.add "Chuang",-19281 
Pinyin.add "Chui",-19275 
Pinyin.add "Chun",-19270 
Pinyin.add "Chuo",-19263 
Pinyin.add "Ci",-19261 
Pinyin.add "Cong",-19249 
Pinyin.add "Cou",-19243 
Pinyin.add "Cu",-19242 
Pinyin.add "Cuan",-19238 
Pinyin.add "Cui",-19235 
Pinyin.add "Cun",-19227 
Pinyin.add "Cuo",-19224 
Pinyin.add "Da",-19218 
Pinyin.add "Dai",-19212 
Pinyin.add "Dan",-19038 
Pinyin.add "Dang",-19023 
Pinyin.add "Dao",-19018 
Pinyin.add "De",-19006 
Pinyin.add "Deng",-19003 
Pinyin.add "Di",-18996 
Pinyin.add "Dian",-18977 
Pinyin.add "Diao",-18961 
Pinyin.add "Die",-18952 
Pinyin.add "Ding",-18783 
Pinyin.add "Diu",-18774 
Pinyin.add "Dong",-18773 
Pinyin.add "Dou",-18763 
Pinyin.add "Du",-18756 
Pinyin.add "Duan",-18741 
Pinyin.add "Dui",-18735 
Pinyin.add "Dun",-18731 
Pinyin.add "Duo",-18722 
Pinyin.add "E",-18710 
Pinyin.add "En",-18697 
Pinyin.add "Er",-18696 
Pinyin.add "Fa",-18526 
Pinyin.add "Fan",-18518 
Pinyin.add "Fang",-18501 
Pinyin.add "Fei",-18490 
Pinyin.add "Fen",-18478 
Pinyin.add "Feng",-18463 
Pinyin.add "Fo",-18448 
Pinyin.add "Fou",-18447 
Pinyin.add "Fu",-18446 
Pinyin.add "Ga",-18239 
Pinyin.add "Gai",-18237 
Pinyin.add "Gan",-18231 
Pinyin.add "Gang",-18220 
Pinyin.add "Gao",-18211 
Pinyin.add "Ge",-18201 
Pinyin.add "Gei",-18184 
Pinyin.add "Gen",-18183 
Pinyin.add "Geng",-18181 
Pinyin.add "Gong",-18012 
Pinyin.add "Gou",-17997 
Pinyin.add "Gu",-17988 
Pinyin.add "Gua",-17970 
Pinyin.add "Guai",-17964 
Pinyin.add "Guan",-17961 
Pinyin.add "Guang",-17950 
Pinyin.add "Gui",-17947 
Pinyin.add "Gun",-17931 
Pinyin.add "Guo",-17928 
Pinyin.add "Ha",-17922 
Pinyin.add "Hai",-17759 
Pinyin.add "Han",-17752 
Pinyin.add "Hang",-17733 
Pinyin.add "Hao",-17730 
Pinyin.add "He",-17721 
Pinyin.add "Hei",-17703 
Pinyin.add "Hen",-17701 
Pinyin.add "Heng",-17697 
Pinyin.add "Hong",-17692 
Pinyin.add "Hou",-17683 
Pinyin.add "Hu",-17676 
Pinyin.add "Hua",-17496 
Pinyin.add "Huai",-17487 
Pinyin.add "Huan",-17482 
Pinyin.add "Huang",-17468 
Pinyin.add "Hui",-17454 
Pinyin.add "Hun",-17433 
Pinyin.add "Huo",-17427 
Pinyin.add "Ji",-17417 
Pinyin.add "Jia",-17202 
Pinyin.add "Jian",-17185 
Pinyin.add "Jiang",-16983 
Pinyin.add "Jiao",-16970 
Pinyin.add "Jie",-16942 
Pinyin.add "Jin",-16915 
Pinyin.add "Jing",-16733 
Pinyin.add "Jiong",-16708 
Pinyin.add "Jiu",-16706 
Pinyin.add "Ju",-16689 
Pinyin.add "Juan",-16664 
Pinyin.add "Jue",-16657 
Pinyin.add "Jun",-16647 
Pinyin.add "Ka",-16474 
Pinyin.add "Kai",-16470 
Pinyin.add "Kan",-16465 
Pinyin.add "Kang",-16459 
Pinyin.add "Kao",-16452 
Pinyin.add "Ke",-16448 
Pinyin.add "Ken",-16433 
Pinyin.add "Keng",-16429 
Pinyin.add "Kong",-16427 
Pinyin.add "Kou",-16423 
Pinyin.add "Ku",-16419 
Pinyin.add "Kua",-16412 
Pinyin.add "Kuai",-16407 
Pinyin.add "Kuan",-16403 
Pinyin.add "Kuang",-16401 
Pinyin.add "Kui",-16393 
Pinyin.add "Kun",-16220 
Pinyin.add "Kuo",-16216 
Pinyin.add "La",-16212 
Pinyin.add "Lai",-16205 
Pinyin.add "Lan",-16202 
Pinyin.add "Lang",-16187 
Pinyin.add "Lao",-16180 
Pinyin.add "Le",-16171 
Pinyin.add "Lei",-16169 
Pinyin.add "Leng",-16158 
Pinyin.add "Li",-16155 
Pinyin.add "Lia",-15959 
Pinyin.add "Lian",-15958 
Pinyin.add "Liang",-15944 
Pinyin.add "Liao",-15933 
Pinyin.add "Lie",-15920 
Pinyin.add "Lin",-15915 
Pinyin.add "Ling",-15903 
Pinyin.add "Liu",-15889 
Pinyin.add "Long",-15878 
Pinyin.add "Lou",-15707 
Pinyin.add "Lu",-15701 
Pinyin.add "Lv",-15681 
Pinyin.add "Luan",-15667 
Pinyin.add "Lue",-15661 
Pinyin.add "Lun",-15659 
Pinyin.add "Luo",-15652 
Pinyin.add "Ma",-15640 
Pinyin.add "Mai",-15631 
Pinyin.add "Man",-15625 
Pinyin.add "Mang",-15454 
Pinyin.add "Mao",-15448 
Pinyin.add "Me",-15436 
Pinyin.add "Mei",-15435 
Pinyin.add "Men",-15419 
Pinyin.add "Meng",-15416 
Pinyin.add "Mi",-15408 
Pinyin.add "Mian",-15394 
Pinyin.add "Miao",-15385 
Pinyin.add "Mie",-15377 
Pinyin.add "Min",-15375 
Pinyin.add "Ming",-15369 
Pinyin.add "Miu",-15363 
Pinyin.add "Mo",-15362 
Pinyin.add "Mou",-15183 
Pinyin.add "Mu",-15180 
Pinyin.add "Na",-15165 
Pinyin.add "Nai",-15158 
Pinyin.add "Nan",-15153 
Pinyin.add "Nang",-15150 
Pinyin.add "Nao",-15149 
Pinyin.add "Ne",-15144 
Pinyin.add "Nei",-15143 
Pinyin.add "Nen",-15141 
Pinyin.add "Neng",-15140 
Pinyin.add "Ni",-15139 
Pinyin.add "Nian",-15128 
Pinyin.add "Niang",-15121 
Pinyin.add "Niao",-15119 
Pinyin.add "Nie",-15117 
Pinyin.add "Nin",-15110 
Pinyin.add "Ning",-15109 
Pinyin.add "Niu",-14941 
Pinyin.add "Nong",-14937 
Pinyin.add "Nu",-14933 
Pinyin.add "Nv",-14930 
Pinyin.add "Nuan",-14929 
Pinyin.add "Nue",-14928 
Pinyin.add "Nuo",-14926 
Pinyin.add "O",-14922 
Pinyin.add "Ou",-14921 
Pinyin.add "Pa",-14914 
Pinyin.add "Pai",-14908 
Pinyin.add "Pan",-14902 
Pinyin.add "Pang",-14894 
Pinyin.add "Pao",-14889 
Pinyin.add "Pei",-14882 
Pinyin.add "Pen",-14873 
Pinyin.add "Peng",-14871 
Pinyin.add "Pi",-14857 
Pinyin.add "Pian",-14678 
Pinyin.add "Piao",-14674 
Pinyin.add "Pie",-14670 
Pinyin.add "Pin",-14668 
Pinyin.add "Ping",-14663 
Pinyin.add "Po",-14654 
Pinyin.add "Pu",-14645 
Pinyin.add "Qi",-14630 
Pinyin.add "Qia",-14594 
Pinyin.add "Qian",-14429 
Pinyin.add "Qiang",-14407 
Pinyin.add "Qiao",-14399 
Pinyin.add "Qie",-14384 
Pinyin.add "Qin",-14379 
Pinyin.add "Qing",-14368 
Pinyin.add "Qiong",-14355 
Pinyin.add "Qiu",-14353 
Pinyin.add "Qu",-14345 
Pinyin.add "Quan",-14170 
Pinyin.add "Que",-14159 
Pinyin.add "Qun",-14151 
Pinyin.add "Ran",-14149 
Pinyin.add "Rang",-14145 
Pinyin.add "Rao",-14140 
Pinyin.add "Re",-14137 
Pinyin.add "Ren",-14135 
Pinyin.add "Reng",-14125 
Pinyin.add "Ri",-14123 
Pinyin.add "Rong",-14122 
Pinyin.add "Rou",-14112 
Pinyin.add "Ru",-14109 
Pinyin.add "Ruan",-14099 
Pinyin.add "Rui",-14097 
Pinyin.add "Run",-14094 
Pinyin.add "Ruo",-14092 
Pinyin.add "Sa",-14090 
Pinyin.add "Sai",-14087 
Pinyin.add "San",-14083 
Pinyin.add "Sang",-13917 
Pinyin.add "Sao",-13914 
Pinyin.add "Se",-13910 
Pinyin.add "Sen",-13907 
Pinyin.add "Seng",-13906 
Pinyin.add "Sha",-13905 
Pinyin.add "Shai",-13896 
Pinyin.add "Shan",-13894 
Pinyin.add "Shang",-13878 
Pinyin.add "Shao",-13870 
Pinyin.add "She",-13859 
Pinyin.add "Shen",-13847 
Pinyin.add "Sheng",-13831 
Pinyin.add "Shi",-13658 
Pinyin.add "Shou",-13611 
Pinyin.add "Shu",-13601 
Pinyin.add "Shua",-13406 
Pinyin.add "Shuai",-13404 
Pinyin.add "Shuan",-13400 
Pinyin.add "Shuang",-13398 
Pinyin.add "Shui",-13395 
Pinyin.add "Shun",-13391 
Pinyin.add "Shuo",-13387 
Pinyin.add "Si",-13383 
Pinyin.add "Song",-13367 
Pinyin.add "Sou",-13359 
Pinyin.add "Su",-13356 
Pinyin.add "Suan",-13343 
Pinyin.add "Sui",-13340 
Pinyin.add "Sun",-13329 
Pinyin.add "Suo",-13326 
Pinyin.add "Ta",-13318 
Pinyin.add "Tai",-13147 
Pinyin.add "Tan",-13138 
Pinyin.add "Tang",-13120 
Pinyin.add "Tao",-13107 
Pinyin.add "Te",-13096 
Pinyin.add "Teng",-13095 
Pinyin.add "Ti",-13091 
Pinyin.add "Tian",-13076 
Pinyin.add "Tiao",-13068 
Pinyin.add "Tie",-13063 
Pinyin.add "Ting",-13060 
Pinyin.add "Tong",-12888 
Pinyin.add "Tou",-12875 
Pinyin.add "Tu",-12871 
Pinyin.add "Tuan",-12860 
Pinyin.add "Tui",-12858 
Pinyin.add "Tun",-12852 
Pinyin.add "Tuo",-12849 
Pinyin.add "Wa",-12838 
Pinyin.add "Wai",-12831 
Pinyin.add "Wan",-12829 
Pinyin.add "Wang",-12812 
Pinyin.add "Wei",-12802 
Pinyin.add "Wen",-12607 
Pinyin.add "Weng",-12597 
Pinyin.add "Wo",-12594 
Pinyin.add "Wu",-12585 
Pinyin.add "Xi",-12556 
Pinyin.add "Xia",-12359 
Pinyin.add "Xian",-12346 
Pinyin.add "Xiang",-12320 
Pinyin.add "Xiao",-12300 
Pinyin.add "Xie",-12120 
Pinyin.add "Xin",-12099 
Pinyin.add "Xing",-12089 
Pinyin.add "Xiong",-12074 
Pinyin.add "Xiu",-12067 
Pinyin.add "Xu",-12058 
Pinyin.add "Xuan",-12039 
Pinyin.add "Xue",-11867 
Pinyin.add "Xun",-11861 
Pinyin.add "Ya",-11847 
Pinyin.add "Yan",-11831 
Pinyin.add "Yang",-11798 
Pinyin.add "Yao",-11781 
Pinyin.add "Ye",-11604 
Pinyin.add "Yi",-11589 
Pinyin.add "Yin",-11536 
Pinyin.add "Ying",-11358 
Pinyin.add "Yo",-11340 
Pinyin.add "Yong",-11339 
Pinyin.add "You",-11324 
Pinyin.add "Yu",-11303 
Pinyin.add "Yuan",-11097 
Pinyin.add "Yue",-11077 
Pinyin.add "Yun",-11067 
Pinyin.add "Za",-11055 
Pinyin.add "Zai",-11052 
Pinyin.add "Zan",-11045 
Pinyin.add "Zang",-11041 
Pinyin.add "Zao",-11038 
Pinyin.add "Ze",-11024 
Pinyin.add "Zei",-11020 
Pinyin.add "Zen",-11019 
Pinyin.add "Zeng",-11018 
Pinyin.add "Zha",-11014 
Pinyin.add "Zhai",-10838 
Pinyin.add "Zhan",-10832 
Pinyin.add "Zhang",-10815 
Pinyin.add "Zhao",-10800 
Pinyin.add "Zhe",-10790 
Pinyin.add "Zhen",-10780 
Pinyin.add "Zheng",-10764 
Pinyin.add "Zhi",-10587 
Pinyin.add "Zhong",-10544 
Pinyin.add "Zhou",-10533 
Pinyin.add "Zhu",-10519 
Pinyin.add "Zhua",-10331 
Pinyin.add "Zhuai",-10329 
Pinyin.add "Zhuan",-10328 
Pinyin.add "Zhuang",-10322 
Pinyin.add "Zhui",-10315 
Pinyin.add "Zhun",-10309 
Pinyin.add "Zhuo",-10307 
Pinyin.add "Zi",-10296 
Pinyin.add "Zong",-10281 
Pinyin.add "Zou",-10274 
Pinyin.add "Zu",-10270 
Pinyin.add "Zuan",-10262 
Pinyin.add "Zui",-10260 
Pinyin.add "Zun",-10256 
Pinyin.add "Zuo",-10254

Function StrToPinyin(AscNum)
	If AscNum>0 and AscNum<160 Then
		StrToPinyin=Chr(AscNum)
	Else
		If AscNum<-20319 or AscNum>-10247 Then
			StrToPinyin=""
		Else
			PinyinId=Pinyin.Items
			PinyinStr=Pinyin.keys
			For i=Pinyin.Count-1 To 0 Step -1
				If PinyinId(i)<=AscNum Then Exit For
			Next
			StrToPinyin=PinyinStr(i)
		End If
	End If
End Function 

Function ToPinyin(ChangeStr)
	ToPinyin=""
	ChangeStr=Replace(ChangeStr,"{","")
	ChangeStr=Replace(ChangeStr,"}","")
	ChangeStr=Replace(ChangeStr,"(","")
	ChangeStr=Replace(ChangeStr,")","")
	ChangeStr=Replace(ChangeStr,"/","")
	ChangeStr=Replace(ChangeStr,"\","")
	ChangeStr=Replace(ChangeStr,"&","")
	ChangeStr=Replace(ChangeStr,"@","")
	ChangeStr=Replace(ChangeStr,",","")
	ChangeStr=Replace(ChangeStr,"<","")
	ChangeStr=Replace(ChangeStr,">","")
	ChangeStr=Replace(ChangeStr,"《","")
	ChangeStr=Replace(ChangeStr,"》","")
	ChangeStr=Replace(ChangeStr,"|","")
	ChangeStr=Replace(ChangeStr,"=","")
	ChangeStr=Replace(ChangeStr,"+","")
	ChangeStr=Replace(ChangeStr," ","")
	ChangeStr=Replace(ChangeStr,"#","")
	ChangeStr=Replace(ChangeStr,"?","")
	For i=1 To len(ChangeStr)
		Select Case Mid(ChangeStr,i,1)
			Case "."
				ToPinyin=Trim(ToPinyin)&"_"
			Case Else
				ToPinyin=Trim(ToPinyin)&Trim(StrToPinyin(Asc(Mid(ChangeStr,i,1))))
		End Select
	Next
End Function 

Response.Write(ToPinyin(Request.QueryString("Str")))
Session.CodePage=65001
Response.Charset = "utf-8"
%>