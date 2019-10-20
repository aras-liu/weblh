<%
'==========================================
'文 件 名：Class/Cls_Fun.asp
'文件用途：常规函数类
'版权所有：
'==========================================

Class Cls_Fun
	Private x,y,ii
	'==============================
	'函 数 名：AlertInfo
	'作    用：错误显示函数
	'参    数：错误提示内容InfoStr，转向页面GoUrl
	'==============================
	Public Function AlertInfo(InfoStr,GoUrl)
		If GoUrl="1" Then
			Response.Write "<Script>alert('"& InfoStr &"');location.href='javascript:history.go(-1)';</Script>"
		Else
			Response.Write "<Script>alert('"& InfoStr &"');location.href='"& GoUrl &"';</Script>"
		End If
		Session.CodePage=936
		Response.End()
	End Function
	
	'==============================
	'函 数 名：HTMLEncode
	'作    用：字符转换函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString=replace(fString, ">", "&gt;")
			fString=replace(fString, "<", "&lt;")
			fString=Replace(fString, CHR(32), " ")		
			fString=Replace(fString, CHR(34), "&quot;")
			fString=Replace(fString, CHR(39), "&#39;")
			fString=Replace(fString, CHR(9), "&nbsp;")
			fString=Replace(fString, CHR(13), "")
			fString=Replace(fString, CHR(10) & CHR(10), "<p></p>")
			fString=Replace(fString, CHR(10), "<br />")
			HTMLEncode=fString
		End If
	End Function
	
	'==============================
	'函 数 名：HTMLDncode
	'作    用：字符转回函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLDncode(fString)
		If Not IsNull(fString) Then
			fString=Replace(fString, "&gt;",">" )
			fString=Replace(fString, "&lt;", "<")
			fString=Replace(fString, " ", CHR(32))
			fString=Replace(fString, "&nbsp;", CHR(9))
			fString=Replace(fString, "&quot;", CHR(34))
			fString=Replace(fString, "&#39;", CHR(39))
			fString=Replace(fString, "", CHR(13))
			fString=Replace(fString, "<p></p>",CHR(10) & CHR(10) )
			fString=Replace(fString, "<br />",CHR(10) )
			HTMLDncode=fString
		End If
	End Function
	
	'==============================
	'函 数 名：AlertNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function AlertNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Call AlertInfo(ErrStr,"1")
		End If
	End Function

	'==============================
	'函 数 名：AlertString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function AlertString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Call AlertInfo(ShortErr,"1")
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Call AlertInfo(LongErr,"1")
		End If
	End Function
	
	'==============================
	'函 数 名：ShowNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function ShowNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Response.Write(ErrStr)
			Call FKDB.DB_Close()
			Session.CodePage=936
			Response.End()
		End If
	End Function

	'==============================
	'函 数 名：ShowString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function ShowString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Response.Write(ShortErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Response.Write(LongErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
	End Function

	'==============================
	'函 数 名：GetCityName
	'作    用：根据城市名称拼音获取城市名
	'参    数：需要城市名称拼音
	'==============================
	Public Function GetCityName(fString)
		Dim chengshi,dangechengshi,fenjie
		chengshi="hengshui,衡水;beijing,北京;tianjin,天津;shanghai,上海;chongqing,重庆;hebei,河北;henan,河南;yunnan,云南;liaoning,辽宁;heilongjiang,黑龙江;hunan,湖南;anhui,安徽;shandong,山东;xinjiang,新疆;jiangsu,江苏;zhejiang,浙江;jiangxi,江西;hubei,湖北;guangxi,广西;gansu,甘肃;shanxi,山西;neimeng,内蒙;shanxi,陕西;jilin,吉林;fujian,福建;guizhou,贵州;guangdong,广东;qinghai,青海;xizang,西藏;sichuan,四川;ningxia,宁夏;hainan,海南;taiwan,台湾;xianggang,香港;aomen,澳门;anping,安平;menggu,蒙古;pengzhou,彭州;yangzhong,扬中;nanchuan,南川;hechuan,合川;yongchuan,永川;aletai,阿勒泰;wusu,乌苏;tacheng,塔城;fukang,阜康;changji,昌吉;kuerle,库尔勒;akesu,阿克苏;hetian,和田;tulufan,吐鲁番;geermu,格尔木;linxia,临夏;hezuo,合作;hancheng,韩城;xingping,兴平;ruili,瑞丽;dali,大理;kaiyuan,开远;gejiu,个旧;simao,思茅;anning,安宁;fuquan,福泉;xingyi,兴义;bijie,毕节;hua,华蓥;emeishan,峨眉山;dujiangyan,都江堰;chongzhou,崇州;qiongshan,琼山;pingxiangshi,凭祥;heshanshi,合山;dongxing,东兴;xi,岑溪;yingde,英德;gaoyao,高要;xinyi,信宜;huazhou,化州;lianjiang,廉江;enping,恩平;lufeng,陆丰;huiyang,惠阳;xingning,兴宁;nanxiong,南雄;chenghai,澄海;conghua,从化;zengcheng,增城;hongjiang,洪江;zixing,资兴;jiang,沅江;linxiang,临湘;changning,常宁;liuyang,浏阳;qianjiang,潜江;enshi,恩施;wuxue,武穴;macheng,麻城;hanchuan,汉川;anlu,安陆;yingcheng,应城;zhijiang,枝江;songzi,松滋;danjiangkou,丹江口;yongcheng,永城;dengzhou,邓州;yima,义马;changge,长葛;yuzhou,禹州;weihui,卫辉;qinyang,沁阳;mengzhou,孟州;wugang,舞钢;yanshi,偃师;xingyang,荥阳;gongyi,巩义;xinzheng,新郑;linqing,临清;leling,乐陵;xintai,新泰;wendeng,文登;rongcheng,荣成;penglai,蓬莱;laizhou,莱州;shouguang,寿光;zhucheng,诸城;qingzhou,青州;anqiu,安丘;tengzhou,滕州;jiaonan,胶南;pingdu,平度;jiaozhou,胶州;zhangqiu,章丘;zhangshu,樟树;fengcheng,丰城;anshan,鞍山;anqing,安庆;anyang,安阳;anshun,安顺;ankang,安康;anguo,安国;aershan,阿尔山;acheng,阿城;anda,安达;anqiu,安丘;anlu,安陆;anning,安宁;akesu,阿克苏;aletai,阿勒泰;anping,安平;beijing,北京;baoding,保定;baotou,包头;benxi,本溪;baishan,白山;baicheng,白城;bangbu,蚌埠;binzhou,滨州;beihai,北海;baise,百色;bazhong,巴中;baoshan,保山;baoji,宝鸡;baiyin,白银;botou,泊头;bazhou,霸州;beining,北宁;beipiao,北票;beian,北安;beiliu,北流;bijie,毕节;bole,博乐;chongqing,重庆;chengde,承德;cangzhou,沧州;changzhi,长治;chifeng,赤峰;chaoyang,朝阳;changchun,长春;changzhou,常州;chuzhou,滁州;chaohu,巢湖;chizhou,池州;changsha,长沙;changde,常德;chenzhou,郴州;chaozhou,潮州;chongzuo,崇左;chengdu,成都;cheng,藁城;chun,珲春;changshu,常熟;cixi,慈溪;changle,长乐;changyi,昌邑;changge,长葛;chibi,赤壁;changning,常宁;conghua,从化;chenghai,澄海;chaoyangshi,潮阳;chongzhou,崇州;chishui,赤水;chuxiong,楚雄;datong,大同;dalian,大连;dandong,丹东;daqing,大庆;dongying,东营;dezhou,德州;dongwan,东莞;deyang,德阳;dazhou,达州;dingzhou,定州;donggang,东港;dashiqiao,大石桥;dengta,灯塔;diaobingshan,调兵山;dehui,德惠;dian,桦甸;daan,大安;dunhua,敦化;dongtai,东台;dafeng,大丰;danyang,丹阳;dongyang,东阳;dexing,德兴;dengfeng,登封;dengzhou,邓州;daye,大冶;danjiangkou,丹江口;dangyang,当阳;dongxing,东兴;dujiangyan,都江堰;duyun,都匀;dali,大理;dunhuang,敦煌;delingha,德令哈;eerduosi,鄂尔多斯;ezhou,鄂州;eerguna,额尔古纳;erlianhaote,二连浩特;enshi,恩施;enping,恩平;emeishan,峨眉山;fushun,抚顺;fuxin,阜新;fuyangshi,阜阳;fuzhou,福州;fuzhoushi,抚州;foshan,佛山;fangchenggang,防城港;fenyang,汾阳;fengzhen,丰镇;fengchengshi,凤城;fujin,富锦;fuyang,富阳;fenghua,奉化;fuqing,福清;fuan,福安;fuding,福鼎;fengcheng,丰城;feicheng,肥城;fuquan,福泉;fukang,阜康;ganzhou,赣州;guangzhou,广州;guilin,桂林;guigang,贵港;guangyuan,广元;guangan,广安;guiyang,贵阳;guyuan,固原;gaobeidian,高碑店;gujiao,古交;gaoping,高平;genhe,根河;gaizhou,盖州;gongzhuling,公主岭;gaoyou,高邮;guixi,贵溪;gaoan,高安;gaomi,高密;gongyi,巩义;guangshui,广水;gaozhou,高州;gaoyao,高要;guiping,桂平;guanghan,广汉;gejiu,个旧;geermu,格尔木;handan,邯郸;hengshui,衡水;huhehaote,呼和浩特;hulunbeier,呼伦贝尔;huludao,葫芦岛;haerbin,哈尔滨;hegang,鹤岗;heihe,黑河;huaian,淮安;hangzhou,杭州;huzhou,湖州;hefei,合肥;huainan,淮南;huaibei,淮北;huangshan,黄山;haozhou,亳州;heze,菏泽;hebi,鹤壁;luohe,漯河;huangshi,黄石;huanggang,黄冈;hengyang,衡阳;huaihua,怀化;heyuan,河源;huizhou,惠州;hezhou,贺州;hechi,河池;haikou,海口;hanzhong,汉中;huang,黄骅;hejian,河间;hejin,河津;houma,侯马;huozhou,霍州;huolinguole,霍林郭勒;haicheng,海城;he,蛟河;helong,和龙;nehe,讷河;hulin,虎林;hailin,海林;hailun,海伦;haimen,海门;haining,海宁;haiyang,海阳;honghu,洪湖;hanchuan,汉川;hongjiang,洪江;huiyang,惠阳;heshan,鹤山;huazhou,化州;heshanshi,合山;hua,华蓥;huayin,华阴;hancheng,韩城;hezuo,合作;hami,哈密;hetian,和田;hechuan,合川;hebei,河北;jincheng,晋城;jinzhong,晋中;jinzhoushi,锦州;jilin,吉林;jixi,鸡西;jiamusi,佳木斯;jiaxing,嘉兴;jinhua,金华;jingdezhen,景德镇;jiujiang,九江;jianshi,吉安;jinan,济南;jining,济宁;jiaozuo,焦作;jingzhou,荆州;jingmen,荆门;jiangmen,江门;jieyang,揭阳;jinchang,金昌;jiayuguan,嘉峪关;jiuquan,酒泉;jinzhou,晋州;jizhou,冀州;jiexiu,介休;jiningshi,集宁;jiutai,九台;jian,集安;jiangdu,江都;jingjiang,靖江;jiangyan,姜堰;jurong,句容;jintan,金坛;jiangyin,江阴;jiande,建德;jiangshan,江山;jieshou,界首;jinjiang,晋江;jianou,建瓯;jianyangshi,建阳;jinggangshan,井冈山;jiaozhou,胶州;jimo,即墨;jiaonan,胶南;jiang,沅江;jishou,吉首;jiangyou,江油;jianyang,简阳;jinghong,景洪;jiangjin,江津;kaifeng,开封;kunming,昆明;kelamayi,克拉玛依;kaiyuan,开原;kunshan,昆山;kaiping,开平;kaili,凯里;kaiyuan,开远;kashi,喀什;kuerle,库尔勒;kuitun,奎屯;langfang,廊坊;linfen,临汾;liaoyang,辽阳;liaoyuan,辽源;lianyungang,连云港;lishui,丽水;liuan,六安;longyan,龙岩;laiwu,莱芜;linyi,临沂;liaocheng,聊城;luoyang,洛阳;loudi,娄底;liuzhou,柳州;laibin,来宾;luzhou,泸州;leshan,乐山;liupanshui,六盘水;lijiang,丽江;lasa,拉萨;lanzhou,兰州;luquan,鹿泉;lucheng,潞城;linhe,临河;linghai,凌海;lingyuan,凌源;linjiang,临江;longjing,龙井;linan,临安;leqing,乐清;lanxi,兰溪;linhai,临海;longquan,龙泉;longhai,龙海;leping,乐平;laixi,莱西;longkou,龙口;laiyang,莱阳;laizhou,莱州;leling,乐陵;linqing,临清;linzhou,林州;lingbao,灵宝;laohekou,老河口;lichuan,利川;ling,醴陵;luo,汨罗;linxiang,临湘;lengshuijiang,冷水江;lianyuan,涟源;lechang,乐昌;lufeng,陆丰;lianjiang,廉江;leizhou,雷州;lianzhou,连州;luoding,罗定;luxi,潞西;linxia,临夏;mudanjiang,牡丹江;maanshan,马鞍山;meizhou,梅州;maoming,茂名;mianyang,绵阳;meishan,眉山;manzhouli,满洲里;meihekou,梅河口;mishan,密山;muleng,穆棱;mingguang,明光;mengzhou,孟州;macheng,麻城;mianzhu,绵竹;miquan,米泉;menggu,蒙古;nanjing,南京;nantong,南通;ningbo,宁波;nanping,南平;ningde,宁德;nanchang,南昌;nanyang,南阳;nanning,南宁;neijiang,内江;nanchong,南充;nangong,南宫;nan,洮南;ningan,宁安;ningguo,宁国;nanan,南安;nankang,南康;nanxiong,南雄;nanchuan,南川;ningxia,宁夏;panjin,盘锦;putian,莆田;pingxiang,萍乡;pingdingshan,平顶山;panzhihua,攀枝花;pingliang,平凉;pulandian,普兰店;panshi,磐石;pizhou,邳州;pinghu,平湖;pingdu,平度;penglai,蓬莱;puning,普宁;pingxiangshi,凭祥;pengzhou,彭州;qinhuangdao,秦皇岛;qiqihaer,齐齐哈尔;qitaihe,七台河;quanzhou,泉州;qingdao,青岛;qingyuan,清远;qinzhou,钦州;qujing,曲靖;qingyang,庆阳;qianan,迁安;qidong,启东;qingzhou,青州;qixia,栖霞;qinyang,沁阳;qianjiang,潜江;qiongshan,琼山;qingzhen,清镇;qingtongxia,青铜峡;rizhao,日照;renqiu,任丘;rugao,如皋;ruian,瑞安;ruichang,瑞昌;ruijin,瑞金;rongcheng,荣成;rushan,乳山;ruzhou,汝州;renhuai,仁怀;ruili,瑞丽;rikaze,日喀则;shanghai,上海;shijiazhuang,石家庄;shuozhou,朔州;shenyang,沈阳;siping,四平;songyuan,松原;shuangyashan,双鸭山;suihua,绥化;suqian,宿迁;suzhou,苏州;shaoxing,绍兴;suzhoushi,宿州;sanming,三明;shangrao,上饶;sanmenxia,三门峡;shangqiu,商丘;shiyan,十堰;suizhou,随州;shaoyang,邵阳;shenzhen,深圳;shantou,汕头;shaoguan,韶关;shanwei,汕尾;sanya,三亚;suining,遂宁;shangluo,商洛;shizuishan,石嘴山;shahe,沙河;sanhe,三河;shenzhou,深州;shulan,舒兰;shuangliao,双辽;shuangcheng,双城;shangzhi,尚志;suifenhe,绥芬河;shangyu,上虞;shengzhou,嵊州;shishi,石狮;shaowu,邵武;shouguang,寿光;yanshi,偃师;shishou,石首;songzi,松滋;shaoshan,韶山;sihui,四会;shifang,什邡;simao,思茅;tianjin,天津;tangshan,唐山;taiyuan,太原;tongliao,通辽;tieling,铁岭;tonghua,通化;taizhoushi,泰州;taizhou,台州;tongling,铜陵;taian,泰安;tongchuan,铜川;tianshui,天水;tumen,图们;tieli,铁力;tongjiang,同江;taixing,泰兴;tongzhou,通州;taicang,太仓;tongxiang,桐乡;tongcheng,桐城;tianchang,天长;tengzhou,滕州;taishan,台山;tongren,铜仁;tulufan,吐鲁番;tacheng,塔城;taiwan,台湾;wuhai,乌海;wuxi,无锡;wenzhou,温州;wuhu,芜湖;weifang,潍坊;weihai,威海;wuhan,武汉;wuzhou,梧州;weinan,渭南;wuwei,武威;wuzhong,吴忠;wulumuqi,乌鲁木齐;wuan,武安;wulanhaote,乌兰浩特;wafangdian,瓦房店;wuchang,五常;wudalianchi,五大连池;wujiang,吴江;wenling,温岭;wuyishan,武夷山;wendeng,文登;wugang,舞钢;weihui,卫辉;wuxue,武穴;wugangshi,武冈;wuchuan,吴川;wanyuan,万源;wusu,乌苏;xingtai,邢台;xinzhou,忻州;xuzhou,徐州;xuzhou,衢州;xuancheng,宣城;xiamen,厦门;xinyu,新余;xinxiang,新乡;xuchang,许昌;xinyang,信阳;xiangfan,襄樊;xiaogan,孝感;xianning,咸宁;xiangtan,湘潭;xian,西安;xianyang,咸阳;xining,西宁;xinji,辛集;xinle,新乐;xiaoyi,孝义;xilinhaote,锡林浩特;xinmin,新民;xingcheng,兴城;xinyishi,新沂;xinghua,兴化;xintai,新泰;xinzheng,新郑;xinmi,新密;xiangcheng,项城;xiangxiang,湘乡;xingning,兴宁;xinyi,信宜;xi,岑溪;xichang,西昌;xingyi,兴义;xuanwei,宣威;xingping,兴平;xizang,西藏;yangquan,阳泉;yuncheng,运城;yingkou,营口;yichun,伊春;yancheng,盐城;yangzhou,扬州;yingtan,鹰潭;yichunshi,宜春;yantai,烟台;puyang,濮阳;yichang,宜昌;yueyang,岳阳;yiyang,益阳;yongzhou,永州;yangjiang,阳江;yunfu,云浮;yulin,玉林;yibin,宜宾;yaan,雅安;yuxi,玉溪;yanan,延安;yulin,榆林;yinchuan,银川;yongji,永济;yuanping,原平;yakeshi,牙克石;yushu,榆树;yanji,延吉;yizheng,仪征;liyang,溧阳;yixing,宜兴;yuyao,余姚;yongkang,永康;yiwu,义乌;yongan,永安;yucheng,禹城;xingyang,荥阳;yuzhou,禹州;yima,义马;yongcheng,永城;yicheng,宜城;yidu,宜都;yingcheng,应城;liuyang,浏阳;leiyang,耒阳;yangchun,阳春;yingde,英德;yizhou,宜州;yumen,玉门;yongchuan,永川;zhangjiakou,张家口;zhenjiang,镇江;zhoushan,舟山;zhangzhou,漳州;zibo,淄博;zaozhuang,枣庄;zhengzhou,郑州;zhoukou,周口;zhumadian,驻马店;zhuzhou,株洲;zhangjiajie,张家界;zhuhai,珠海;zhongshan,中山;zhanjiang,湛江;zhaoqing,肇庆;zigong,自贡;ziyang,资阳;zunyi,遵义;zhaotong,昭通;zhangye,张掖;zunhua,遵化;zhuozhou,涿州;zhalantun,扎兰屯;zhuanghe,庄河;zhaodong,肇东;zhangjiagang,张家港;zhu,诸暨;zhangping,漳平;zhangshu,樟树;zhangqiu,章丘;zhucheng,诸城;zhaoyuan,招远;zaoyang,枣阳;zhijiang,枝江;zhongxiang,钟祥;zixing,资兴;zengcheng,增城;langzhong,阆中;taocheng,桃城区;kaifaqu,开发区;wuyixian,武邑;wuqiang,武强;zaoqiang,枣强;raoyang,饶阳;gucheng,故城;jingxian,景县;fucheng,阜城;jizhou,冀州;shenzhou,深州;daying,大营;neimenggu,内蒙古;zhongwei,中卫;jiaozuo,焦作"
		dangechengshi=Split(chengshi,";")
		GetCityName = ""
		for i=1 to ubound(dangechengshi)+1
			fenjie=Split(dangechengshi(i-1),",")
			If(fString = fenjie(0)) Then
				GetCityName = fenjie(1)
				
			End If
		Next
		
	End Function
	
	'==============================
	'函 数 名：StringLength
	'作    用：判断字符串长度
	'参    数：需进行判断的文本Txt
	'==============================
	Public Function StringLength(Txt)
		Txt=Trim(Txt)
		x=Len(Txt)
		y=0
		For ii=1 To x
			If Asc(Mid(Txt,ii,1))<=2 or Asc(Mid(Txt,ii,1))>255 Then
				y=y + 2
			Else
				y=y + 1
			End If
		Next
		StringLength=y
	End Function
	
	'==============================
	'函 数 名：BeSelect
	'作    用：判断select选项选中
	'参    数：Select1,Select2
	'==============================
	Public Function BeSelect(Select1,Select2)
		If Select1=Select2 Then
			BeSelect=" selected='selected'"
		End If
	End Function
	
	'==============================
	'函 数 名：BeCheck
	'作    用：判断Check选项选中
	'参    数：Check1,Check2
	'==============================
	Public Function BeCheck(Check1,Check2)
		If Check1=Check2 Then
			BeCheck=" checked='checked'"
		End If
	End Function
	
	'==============================
	'函 数 名：CheckModule
	'作    用：判断模块类型，输出名称
	'参    数：要判断的类型ModuleId
	'==============================
	Public Function CheckModule(ModuleId)
		For i=0 To UBound(FKModuleId)
			If ModuleId=Clng(FKModuleId(i)) Then
				CheckModule=FKModuleName(i)
				Exit Function
			End If
		Next
	End Function	

	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCode(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&"1');"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows-1)&"');"">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows+1)&"');"">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&AllPage&"');"">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""SetRContent('MainRight','"&PageUrl&"'+this.options[this.selectedIndex].value);"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function

	'==============================
	'函 数 名：GetNowUrl
	'作    用：返回当前网址
	'参    数：
	'==============================
	Public Function GetNowUrl()
		GetNowUrl=Request.ServerVariables("Script_Name")&"?"&Request.ServerVariables("QUERY_STRING")
	End Function

	'==============================
	'函 数 名：ReplaceTest
	'作    用：正则表达式，替换字符串
	'参    数：规则patrn，要替换的字符串Str，替换为字符串replStr
	'==============================
	Public Function ReplaceTest(patrn,replStr,Str)
		Dim regEx
		Set regEx=New RegExp
		regEx.Pattern=patrn
		regEx.IgnoreCase=True
		regEx.Global=True 
		ReplaceTest=regEx.Replace(Str,replStr)
	End Function 

	'==============================
	'函 数 名：RegExpTest
	'作    用：正则表达式，获取字符串
	'参    数：源字符串patrn，规则strng
	'==============================
	Public Function RegExpTest(patrn, strng)
		Dim regEx, Matchs, Matches, RetStr
		Set regEx=New RegExp 
		regEx.Pattern=patrn 
		regEx.IgnoreCase=True
		regEx.Global=True
		Set Matches=regEx.Execute(strng) 
		For Each Matchs in Matches
			RetStr=RetStr & Matchs.Value & "|-_-|"
		Next 
		RegExpTest=RetStr 
	End Function 
	
	'==============================
	'函 数 名：NoTrash
	'作    用：垃圾信息强力判断
	'参    数：
	'TryStr      要判断的信息
	'CheckType   判断类型
	'ErrStr      垃圾信息提示
	'==============================
	Public Function NoTrash(TryStr,CheckType,ErrStr)
		Dim HttpCount,TryStr2,TryStr3
		TryStr2=LCase(TryStr)
		TryStr3=Replace(TryStr2,"|-^-|fangka|-^-|","")
		TryStr3=Replace(TryStr3,"|-_-|fangka|-_-|","")
		HttpCount=Clng(((Len(TryStr2)-Len(Replace(TryStr2,"http://","")))/7))
		If HttpCount>3 Then
			Call AlertInfo(ErrStr,"1")
		End If
		If CheckType=1 Then
			If StringLength(TryStr3)<=(Len(TryStr3)*1.2) Then
				Call AlertInfo(ErrStr,"1")
			End If
		End If
	End Function

	'==============================
	'函 数 名：GetHttpPage
	'作    用：获取页面源代码函数
	'参    数：网址HttpUrl，编码Cset
	'==============================
	Public Function GetHttpPage(HttpUrl,Cset)
		If IsNull(HttpUrl)=True Or HttpUrl="" Then
			GetHttpPage2="ERR！"
			Exit Function
		End If
		On Error Resume Next
		Dim Http
		Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
		Http.open "GET",HttpUrl,False
		Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing
			GetHttpPage="ERR！"
			Exit function
		End if
		GetHttpPage2=BytesToBSTR(Http.responseBody,Cset)
		Set Http=Nothing
		If Err.number<>0 then
			Err.Clear
			GetHttpPage2="ERR！"
			Exit function
		End If
	End Function

	'==============================
	'函 数 名：BytesToBstr
	'作    用：转换编码函数
	'参    数：字符串Body，编码Cset
	'==============================
	Private Function BytesToBstr(Body,Cset)
		Dim Objstream
		Set Objstream=Server.CreateObject("ado"&"d"&"b.st"&"re"&"am")
		Objstream.Type=1
		Objstream.Mode =3
		Objstream.Open
		Objstream.Write body
		Objstream.Position=0
		Objstream.Type=2
		Objstream.Charset=Cset
		BytesToBstr=Objstream.ReadText 
		Objstream.Close
		set Objstream=nothing
	End Function
	
	'==============================
	'函 数 名：HtmlToJs
	'作    用：HTML转JS
	'参    数：字符串CStrs
	'==============================
	Public Function HtmlToJs(CStrs)
		Dim ToJs
		CStrs=Replace(CStrs,Chr(10),"") 
		CStrs=Replace(CStrs,Chr(32)&Chr(32),"") 
		CStrs=Split(CStrs,Chr(13))
		ToJs=""
		For i=0 To UBound(CStrs) 
		If Trim(CStrs(i)) <> "" Then 
			CStrs(i)= Replace(CStrs(i),Chr(34),Chr(39)) 
			ToJs=ToJs&"document.write("&Chr(34)&CStrs(i)&Chr(34)&");"&Chr(10) 
		End If 
		Next
		HtmlToJs=ToJs
	End Function
		
	'==============================
	'函 数 名：UnEscape
	'作    用：js escape解码
	'参    数：字符串Str
	'==============================
	Public Function UnEscape(Str)
		dim r,s,c 
		s="" 
		For r=1 to Len(Str) 
			c=Mid(Str,r,1) 
			If Mid(Str,r,2)="%u" and r<=Len(Str)-5 Then 
				If IsNumeric("&H" & Mid(Str,r+2,4)) Then 
					s=s & CHRW(CInt("&H" & Mid(Str,r+2,4))) 
					r=r+5 
				Else 
					s=s & c 
				End If 
			ElseIf c="%" and r<=Len(Str)-2 Then 
				If IsNumeric("&H" & Mid(Str,r+1,2)) Then 
					s=s & CHRW(CInt("&H" & Mid(Str,r+1,2))) 
					r=r+2 
				Else 
					s=s & c 
				End If 
			Else 
				s=s & c 
			End If 
		Next 
		UnEscape=s 
	End Function 
	
	'==============================
	'函 数 名：RemoveHTML
	'作    用：过滤HTML
	'参    数：
	'==============================
	Public Function RemoveHTML(strHTML)
		Dim objRegExp, Match, Matches 
		Set objRegExp=New Regexp 
		objRegExp.IgnoreCase=True 
		objRegExp.Global=True 
		'取闭合的<> 
		objRegExp.Pattern="<.+?>" 
		'进行匹配 
		Set Matches=objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		'取特殊字符
		objRegExp.Pattern="\&.+?;" 
		'进行匹配 
		Set Matches=objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		RemoveHTML=strHTML 
		Set objRegExp=Nothing 
	End Function
	
	'==============================
	'函 数 名：IsObjInstalled
	'作    用：判断组件是否安装了
	'参    数：组件名strClassString
	'==============================
	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled=False 
		Err.Clear
		Dim xTestObj
		Set xTestObj=Server.CreateObject(strClassString)
		If 0=Err Then IsObjInstalled=True 
		Set xTestObj=Nothing
		Err.Clear
	End Function
	
	'==============================
	'函 数 名：ReplaceExt
	'作    用：过滤敏感后缀
	'参    数：要替换的字符串ReStr
	'==============================
	Public Function ReplaceExt(ReStr)
		ReStr=Replace(ReStr,"aspx","")
		ReStr=Replace(ReStr,"asp","")
		ReStr=Replace(ReStr,"jsp","")
		ReStr=Replace(ReStr,"cer","")
		ReStr=Replace(ReStr,"asa","")
		ReStr=Replace(ReStr,"cgi","")
		ReplaceExt=ReStr
	End Function
	
	'==============================
	'函 数 名：GetNumeric
	'作    用：获取处理数字参数
	'参    数：
	'GetTag      获取的值
	'GetDefault  默认值
	'==============================
	Public Function GetNumeric(GetTag,GetDefault)
		GetNumeric=Trim(Request.QueryString(GetTag))
		If GetNumeric<>"" Then
			GetNumeric=Clng(GetNumeric)
		Else
			GetNumeric=GetDefault
		End If
	End Function
	
	'==============================
	'函 数 名：ShowErr
	'作    用：错误提示函数
	'参    数：错误说明ErrStr
	'         错误类型ErrType(0前台，1后台)
	'==============================
	Public Function ShowErr(ErrStr,ErrType)
		If ErrType=0 Then
			Temp="<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
			Temp=Temp&"<html xmlns=""http://www.w3.org/1999/xhtml"">"
			Temp=Temp&"<head>"
			Temp=Temp&"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
			Temp=Temp&"<title>系统提示</title>"
			Temp=Temp&"<style type=""text/css"">"
			Temp=Temp&"body,td,th {font-size: 12px;}"
			Temp=Temp&"body {margin-left:0px;margin-top:50px;margin-right: 0px;margin-bottom: 0px;}"
			Temp=Temp&"#errBox {width:500px;margin:0 auto;border:1px solid #CCC;}"
			Temp=Temp&"#errTitle {border-bottom:1px solid #CCC;line-height:25px;text-align:center;}"
			Temp=Temp&"#errStr {padding:10px;line-height:22px;text-indent:24px;}"
			Temp=Temp&"</style>"
			Temp=Temp&"</head>"
			Temp=Temp&"<body>"
			Temp=Temp&"<div id=""errBox"">"
			Temp=Temp&"	<div id=""errTitle"">系统提示</div>"
			Temp=Temp&"    <div id=""errStr"">"&ErrStr&"</div>"
			Temp=Temp&"</div>"
			Temp=Temp&"</body>"
			Temp=Temp&"</html>"
		End If
		If ErrType=1 Then
			Temp="<div id=""errBox"">"
			Temp=Temp&"	<div id=""errTitle"">系统提示</div>"
			Temp=Temp&"    <div id=""errStr"">"&ErrStr&"</div>"
			Temp=Temp&"</div>"
			Temp=Temp&"</body>"
		End If
		If ErrType=2 Then
			Temp=ErrStr
		End If
		Response.Write(Temp)
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End Function
	
	'==============================
	'函 数 名：Jmail
	'作    用：邮件发送
	'参    数：发送到mailTo，邮件主题mailTopic，邮件内容mailBody，编码mailCharset，正文格式mailContentType
	'==============================
	Public Function Jmail(mailTo,mailTopic,mailBody,mailCharset,mailContentType) 
		If IsObjInstalled("JMail.Message") Then
			Dim myJmail 
			Set myJmail=Server.CreateObject("JMail.Message") 
			myJmail.Charset=mailCharset
			myJmail.silent=true 
			myJmail.ContentType=mailContentType
			myJmail.MailServerUserName=Mail_Name 
			myJmail.MailServerPassWord=Mail_Pass 
			myJmail.AddRecipient mailTo,""
			myJmail.Subject=mailTopic 
			myJmail.Body=mailBody
			myJmail.FromName="" 
			myJmail.From=Mail_Address 
			myJmail.Priority=3
			myJmail.Send(Mail_Smtp) 
			myJmail.Close 
			Set myJmail=nothing 
			If Err Then 
				Jmail=Err.Description 
				Err.Clear 
			Else 
				Jmail="发送成功" 
			End If 
		Else
			Jmail="不支持JMAIL组件"
		End If
	End Function 
End Class
%>
