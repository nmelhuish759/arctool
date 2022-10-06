/***** PARSER FOR XLSX *****/
var XParser = {
	styles:{CellXf:[],Fills:[],NumberFmt:[],Fonts:[],Borders:[]},
	themes: {},
	guess: {s: {r:2000000, c:2000000}, e: {r:0, c:0} },
	opts: {WTF:false, bookDeps:false, bookFiles:false, bookProps:false, bookSheets:false, bookVBA:false, cellDates:true, cellFormula:true, cellHTML:true, cellNF:false, cellStyles:true, cellText:true, password:"", sheetRows:0, sheetStubs:false, type:"buffer"},			
	tagregex: new RegExp(/<[\/\?]?[a-zA-Z0-9:]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s?[\/\?]?>/g),
	encregex: new RegExp(/&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/g),
	coderegex: new RegExp(/_x([\da-fA-F]{4})_/g),
	crefregex: new RegExp(/(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g),
	refregex: new RegExp(/ref=["']([^"']*)["']/),
	rregex: new RegExp(/<(?:\w+:)?r>/g),
	rend: new RegExp(/<\/(?:\w+:)?r>/), 
	nlregex: new RegExp(/\r\n/g),
	sitregex: new RegExp(/<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g), 
	sirregex: new RegExp(/<(?:\w+:)?r>/),
	sirphregex: new RegExp(/<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g),
	sheetdataregex: new RegExp(/<(?:\w+:)?sheetData>([\s\S]*)<\/(?:\w+:)?sheetData>/),
	cellXfRegex: new RegExp(/<(?:\w+:)?cellXfs([^>]*)>[\S\s]*?<\/(?:\w+:)?cellXfs>/),
	rowregex: new RegExp(/<\/(?:\w+:)?row>/),
	cellregex: new RegExp(/<(?:\w+:)?c[ >]/),
	fillsRegex: new RegExp(/<(?:\w+:)?fills([^>]*)>[\S\s]*?<\/(?:\w+:)?fills>/),
	numFmtRegex: new RegExp(/<(?:\w+:)?numFmts([^>]*)>[\S\s]*?<\/(?:\w+:)?numFmts>/),
	fontsRegex: new RegExp(/<(?:\w+:)?fonts([^>]*)>[\S\s]*?<\/(?:\w+:)?fonts>/),
	bordersRegex: new RegExp(/<(?:\w+:)?borders([^>]*)>[\S\s]*?<\/(?:\w+:)?borders>/),
	decregex: /[&<>'"]/g,
	htmlcharegex: /[\u0000-\u001f]/g,
	_ssfopts: {date1904: false},
	RBErr: {'#DIV/0!': 7, '#GETTING_DATA': 43, '#N/A': 42, '#NAME?': 29,	'#NULL!': 0, '#NUM!': 36, '#REF!': 23, '#VALUE!': 15, '#WTF?': 255},
	gotCellXfAndTheme: {},
	parsedData: "",
	encodings:{
		'&quot;': '"',
		'&apos;': "'",
		'&gt;': '>',
		'&lt;': '<',
		'&amp;': '&'
	},
	SSFImplicit: {
		"5": '"$"#,##0_);\\("$"#,##0\\)',
		"6": '"$"#,##0_);[Red]\\("$"#,##0\\)',
		"7": '"$"#,##0.00_);\\("$"#,##0.00\\)',
		"8": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		"23": 'General', "24": 'General', "25": 'General', "26": 'General',
		"27": 'm/d/yy', "28": 'm/d/yy', "29": 'm/d/yy', "30": 'm/d/yy', "31": 'm/d/yy',
		"32": 'h:mm:ss', "33": 'h:mm:ss', "34": 'h:mm:ss', "35": 'h:mm:ss',
		"36": 'm/d/yy',
		"41": '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
		"42": '_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
		"43": '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
		"44": '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
		"50": 'm/d/yy', "51": 'm/d/yy', "52": 'm/d/yy', "53": 'm/d/yy', "54": 'm/d/yy',
		"55": 'm/d/yy', "56": 'm/d/yy', "57": 'm/d/yy', "58": 'm/d/yy',
		"59": '0',
		"60": '0.00',
		"61": '#,##0',
		"62": '#,##0.00',
		"63": '"$"#,##0_);\\("$"#,##0\\)',
		"64": '"$"#,##0_);[Red]\\("$"#,##0\\)',
		"65": '"$"#,##0.00_);\\("$"#,##0.00\\)',
		"66": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		"67": '0%',
		"68": '0.00%',
		"69": '# ?/?',
		"70": '# ??/??',
		"71": 'm/d/yy',
		"72": 'm/d/yy',
		"73": 'd-mmm-yy',
		"74": 'd-mmm',
		"75": 'mmm-yy',
		"76": 'h:mm',
		"77": 'h:mm:ss',
		"78": 'm/d/yy h:mm',
		"79": 'mm:ss',
		"80": '[h]:mm:ss',
		"81": 'mmss.0'
	},
	
	getThemeAndStyle: function(zipObj,callback){
		if(!zipObj.hasTheme){
			var styleObj = {}
			var themeData = zipObj.themeData
			var styleData = zipObj.styleData
			styleObj.themes = XParser.parse_theme_xml(themeData,XParser.opts) //create theme obj
			styleObj.styles = {CellXf:[],Fills:[],NumberFmt:[],Fonts:[],Borders:[]} //resetting if there are multiple excel sheets
			var cellXf = styleData.match(XParser.cellXfRegex)
			var fills = styleData.match(XParser.fillsRegex)
			var numFmt = styleData.match(XParser.numFmtRegex)
			var fonts = styleData.match(XParser.fontsRegex)
			var border = styleData.match(XParser.bordersRegex)
			if(cellXf && cellXf[0])
				styleObj.styles = XParser.parse_cellXfs(cellXf[0],styleObj) //create styles.CellXf       Add error handling here
			if(fills && fills[0])
				styleObj.styles = XParser.parse_fills(fills[0],styleObj)
			if(numFmt && numFmt[0])
				styleObj.styles = XParser.parse_numFmts(numFmt[0],styleObj)
			if(fonts && fonts[0])
				styleObj.styles = XParser.parse_fonts(fonts[0],styleObj)
			if(border && border[0])
				styleObj.styles = XParser.parse_borders(border[0],styleObj)

			zipObj.hasTheme = true
			callback(styleObj)			
		
		}else{
			callback(zipObj.styleObj)
		}
		
	},

	parseData: function(data,callback){
		XParser.parsedData = "" //clear it out
		//XParser.getThemeAndStyle(function(){
			//if the data is small then the sheetregex will hit, but if the data is large and it is chunking, then it wont
		//	if(XParser.sheetdataregex.test(data)){
				//var rowData = data.match(XParser.sheetdataregex)
				//rowData = rowData[1]
				//XParser.parse_xml_from_excel(rowData)
			//}else{
				XParser.parse_xml_from_excel(data)
			//}
			return XParser.parsedData
		//})
	},
	//parse_xml_from_excel: function(data, s, XParser.opts, XParser.guess, themes){
	parse_xml_from_excel: function(data,zipObj){
		XParser.getThemeAndStyle(zipObj, function(styleObj){
			zipObj.styleObj = styleObj
		
			XParser.parsedData = "" 
			if(styleObj.fileTypeStrings){
				callback(data.replace( /(<([^>]+)>)/ig, ' '))
				//callback(data)
				return;
			}
			var ri = 0, x = "", cells = [], cref = [], idx=0, i=0, cc=0, d="", p;
			var tag, tagr = 0, tagc = 0;
			var match_f = matchtag("f");
			var match_v = matchtag("v")
			var sstr, ftag;
			var fmtid = 0, fillid = 0;
			var do_format = Array.isArray(styleObj.styles.CellXf), cf;
			var arrayf = [];
			var sharedf = [];
			//var dense = Array.isArray(s);
			var rows = [], rowobj = {}, rowrite = false;
	/*		var marr = data.split(XParser.rowregex)
			//var marr = XParser.split(data,["<","/","r","o","w",">"]) //wrote my own splitter because its faster
			var marrlen = marr.length
			//for(var marr = data.split(XParser.rowregex), mt = 0, marrlen = marr.length; mt != marrlen; ++mt) {
			//for(var mt = marrlen; mt--;) {
			while(marrlen--){
				x = marr[marrlen].trim();
				//var xlen = x.length;
				//if(xlen === 0) continue;
				if(x.indexOf("<v>") == -1) //if empty row
					continue;
	*/	
				/*
				for(ri = 0; ri < xlen; ++ri) if(x.charCodeAt(ri) === 62) break; ++ri;
				tag = XParser.parsexmltag(x.slice(0,ri), true);
				tagr = tag.r != null ? parseInt(tag.r, 10) : tagr+1; tagc = -1;
				if(XParser.opts.sheetRows && XParser.opts.sheetRows < tagr) continue;
				if(XParser.guess.s.r > tagr - 1) XParser.guess.s.r = tagr - 1;
				if(XParser.guess.e.r < tagr - 1) XParser.guess.e.r = tagr - 1;

				if(XParser.opts && XParser.opts.cellStyles) {
					rowobj = {}; rowrite = false;
					if(tag.ht) { rowrite = true; rowobj.hpt = parseFloat(tag.ht); rowobj.hpx = XParser.pt2px(rowobj.hpt); }
					if(tag.hidden == "1") { rowrite = true; rowobj.hidden = true; }
					if(tag.outlineLevel != null) { rowrite = true; rowobj.level = +tag.outlineLevel; }
					if(rowrite) rows[tagr-1] = rowobj;
				}
				*/
				
				//cells = x.slice(ri).split(XParser.cellregex);
				cells = data.split(XParser.cellregex);
				//cells = XParser.split(data,['<','/','c','>'])
				//cells = XParser.indexOf(data,"<c ","</c>")
				var cellLen = cells.length
				var ri = 0;
				//for(ri = 0; ri != cells.length; ++ri) {
				for(; ri < cellLen; ri++) {
					x = cells[ri].trim();
					//if(x.length === 0) continue;
					if((x.indexOf("<v>") == -1) || (x.indexOf('t="s"') > -1)){ //if no value or if string
						continue;
					}
					//cref = x.match(XParser.rregex); 
					if(XParser.rregex.test(x)) 
						cref = x.match(XParser.rregex); 
					idx = ri; 
					i=0; 
					cc=0;
					x = "<c " + (x.slice(0,1)=="<"?">":"") + x;
					/*if(cref != null && cref.length === 2) {
						idx = 0; d=cref[1];
						for(i=0; i != d.length; ++i) {
							if((cc=d.charCodeAt(i)-64) < 1 || cc > 26) break;
							idx = 26*idx + cc;
						}
						--idx;
						tagc = idx;
					} else ++tagc;
					for(i = 0; i != x.length; ++i) if(x.charCodeAt(i) === 62) break; ++i;
					tag = parsexmltag(x.slice(0,i), true);
					*/
					//var splitValueOff = x.split(">")[0]+">"
					/*var splitValueOff = XParser.split(x,[">"])[0]+">"
					tag = XParser.parsexmltag(splitValueOff, true);*/
					++tagc
					tag = XParser.parsexmltag(x, true);
					if(!tag.r) tag.r = XParser.encode_cell({r:tagr-1, c:tagc});
					d = x.slice(i);
					p = ({t:""});
					if((cref = match_v.exec(d))!= null && cref[1] !== '') //im gonna use exec here since it is only getting the first regex found and its faster
						//p.v = XParser.unescapexml(cref[1]); originial might still need? 
						p.v = cref[1]
						
			/*		if(XParser.opts.cellFormula) {
						if((cref = d.match(match_f))!= null && cref[1] !== '') {

							p.f = XParser._xlfn(XParser.unescapexml(XParser.utf8read(cref[1])));
							if(cref[0].indexOf('t="array"') > -1) {
								p.F = (d.match(XParser.refregex)||[])[1];
								if(p.F.indexOf(":") > -1) arrayf.push([XParser.safe_decode_range(p.F), p.F]);
							} else if(cref[0].indexOf('t="shared"') > -1) {

								ftag = XParser.parsexmltag(cref[0]);
								sharedf[parseInt(ftag.si, 10)] = [ftag, XParser._xlfn(XParser.unescapexml(XParser.utf8read(cref[1])))];
							}
						} else if((cref = d.match(/<f[^>]*\/>/))) {
							ftag = XParser.parsexmltag(cref[0]);
							if(sharedf[ftag.si]) p.f = XParser.shift_formula_xlsx(sharedf[ftag.si][1], sharedf[ftag.si][0].ref, tag.r);
						}
		
						var _tag = XParser.decode_cell(tag.r);
						for(i = 0; i < arrayf.length; ++i)
							if(_tag.r >= arrayf[i][0].s.r && _tag.r <= arrayf[i][0].e.r)
								if(_tag.c >= arrayf[i][0].s.c && _tag.c <= arrayf[i][0].e.c)
									p.F = arrayf[i][1];
					}
				*/	
					if(tag.t == null && p.v === undefined) {
						if(p.f || p.F) {
							p.v = 0; p.t = "n";
						} else if(!XParser.opts.sheetStubs){ 
							continue; 
						}else{
							p.t = "z";
						}
					}
					else p.t = tag.t || "n";
					if(XParser.guess.s.c > tagc) XParser.guess.s.c = tagc;
					if(XParser.guess.e.c < tagc) XParser.guess.e.c = tagc;
					/* 18.18.11 t ST_CellType */
					switch(p.t) {
						case 'n':
							if(p.v == "" || p.v == null) {
								if(!XParser.opts.sheetStubs){
									continue; 
								}
								p.t = 'z';
							} else p.v = parseFloat(p.v);
							break;
						case 's':
							//p.v = ""
							/*if(typeof p.v == 'undefined') {
								if(!XParser.opts.sheetStubs) continue;
								p.t = 'z';
							} else {
								sstr = strs[parseInt(p.v, 10)];
								p.v = sstr.t;
								p.r = sstr.r;
								if(XParser.opts.cellHTML) p.h = sstr.h;
							}*/
							//not sure if i want to do anything here
							break;
						case 'str':
							p.t = "s";
							p.v = (p.v!=null) ? XParser.utf8read(p.v) : '';
							if(XParser.opts.cellHTML) p.h = XParser.escapehtml(p.v);
							break;
						case 'inlineStr':
							cref = d.match(isregex);
							p.t = 's';
							if(cref != null && (sstr = parse_si(cref[1]))) p.v = sstr.t; else p.v = "";
							break;
						case 'b': p.v = XParser.parsexmlbool(p.v); break;
						case 'd':
							if(XParser.opts.cellDates) p.v = XParser.parseDate(p.v, 1);
							else { p.v = XParser.datenum(parseDate(p.v, 1)); p.t = 'n'; }
							break;
						/* error string in .w, number in .v */
						case 'e':
							if(!XParser.opts || XParser.opts.cellText !== false) p.w = p.v;
							p.v = XParser.RBErr[p.v]; break;
					}
					/* formatting */
					fmtid = fillid = 0;
					if(do_format && tag.s !== undefined) {
						cf = styleObj.styles.CellXf[tag.s];
						if(cf != null) {
							if(cf.numFmtId != null) fmtid = cf.numFmtId;
							if(XParser.opts.cellStyles) {
								if(cf.fillId != null) fillid = cf.fillId;
							}
						}
					}
					XParser.safe_format(p, fmtid, fillid, XParser.opts,styleObj);
					if(XParser.opts.cellDates && do_format && p.t == 'n' && SSF.is_date(SSF._table[fmtid])) { p.t = 'd'; p.v = XParser.numdate(p.v); }
					XParser.parsedData += p.w
					XParser.parsedData += "  " //add two spaces so numbers like 1234 5678 9123 4567 dont get picked up as credit card numbers 
					/*if(dense) {
						var _r = decode_cell(tag.r);
						if(!s[_r.r]) s[_r.r] = [];
						s[_r.r][_r.c] = p;
					} else s[tag.r] = p;*/
				}
			//} //while loop
			//return XParser.parsedData
			postMessage({"data":XParser.parsedData, "zip":zipObj})
		})
	
	},
	
	indexOf: function(data,startChar,endChar){
		
		var startCharLen = startChar.length
		var endCharLen = endChar.length
		var start = data.indexOf(startChar)
		var arr = []
		var end = data.indexOf(endChar)
		arr.push(data.slice(start,end))

		while(start > -1){
			start = data.indexOf(startChar,start+startCharLen)
			end = data.indexOf(endChar,start+endCharLen)
			arr.push(data.slice(start,end))
		}
		return arr		
	},
	
	split: function(data,findWord){
		//var findWord = ["<","/","r","o","w",">"]
		var foundWordIndex = 0
		var foundWord = ""
		var startIndex = 0
		var endIndex = 0
		var arr = []
		var i = 0;
		var len = data.length
		var findLen = findWord.length
		var findIndex = findWord[foundWordIndex]
		var foundWordMax = true
		for(; i < len; i++){
			endIndex++			
			if((findIndex == data[i]) && foundWordMax){
				foundWordIndex++
				findIndex = findWord[foundWordIndex]
				foundWordMax = foundWordIndex <= findLen ? true : false
			}else if(foundWordIndex == findLen){
				arr.push(data.slice(startIndex,endIndex))
				startIndex = endIndex
				foundWord = ""
				foundWordIndex = 0
				findIndex = findWord[foundWordIndex]
				foundWordMax = true
			}else{
				foundWordIndex = 0
				findIndex = findWord[foundWordIndex]
			}
		}
		arr.push(data.slice(startIndex,endIndex))
		return arr
	},
	
	safe_format: function(p, fmtid, fillid, opts, styleObj) {
		var themes = styleObj.themes
		var styles = styleObj.styles
		if(p.t === 'z') return;
		if(p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
		try {
			if(opts.cellNF) p.z = SSF._table[fmtid];
		} catch(e) { if(opts.WTF) throw e; }
		if(!opts || opts.cellText !== false) try {
			if(SSF._table[fmtid] == null) SSF.load(XParser.SSFImplicit[fmtid] || "General", fmtid);
			if(p.t === 'e') p.w = p.w || BErr[p.v];
			else if(fmtid === 0) {
				if(p.t === 'n') {
					if((p.v|0) === p.v) p.w = SSF._general_int(p.v);
					else p.w = SSF._general_num(p.v);
				}
				else if(p.t === 'd') {
					var dd = XParser.datenum(p.v);
					if((dd|0) === dd) p.w = SSF._general_int(dd);
					else p.w = SSF._general_num(dd);
				}
				else if(p.v === undefined) return "";
				else p.w = SSF._general(p.v,XParser._ssfopts);
			}
			else if(p.t === 'd') p.w = SSF.format(fmtid,XParser.datenum(p.v),XParser._ssfopts);
			else p.w = SSF.format(fmtid,p.v,XParser._ssfopts);
		} catch(e) { if(opts.WTF) throw e; }
		if(!opts.cellStyles) return;
		if(fillid != null) try {
			p.s = styles.Fills[fillid];
			if (p.s.fgColor && p.s.fgColor.theme && !p.s.fgColor.rgb) {
				p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0);
				if(opts.WTF) p.s.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fgColor.theme].rgb;
			}
			if (p.s.bgColor && p.s.bgColor.theme) {
				p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0);
				if(opts.WTF) p.s.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.bgColor.theme].rgb;
			}
		} catch(e) { if(opts.WTF && styles.Fills) throw e; }
	},

	
	parse_cellXfs: function(str,styleObj) {
		var xf;
		//var xf = XParser.xf;
		var pass = false;
		var cellXF_uint = [ "numFmtId", "fillId", "fontId", "borderId", "xfId" ];
		var cellXF_bool = [ "applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix" ];
		
		str.match(XParser.tagregex).forEach(function(x) {
			var y = XParser.parsexmltag(x), i = 0;
			switch(XParser.strip_ns(y[0])) {
				case '<cellXfs': case '<cellXfs>': case '<cellXfs/>': case '</cellXfs>': break;

				/* 18.8.45 xf CT_Xf */
				case '<xf': case '<xf/>':
					xf = y;
					delete xf[0];
					for(i = 0; i < cellXF_uint.length; ++i) if(xf[cellXF_uint[i]])
						xf[cellXF_uint[i]] = parseInt(xf[cellXF_uint[i]], 10);
					for(i = 0; i < cellXF_bool.length; ++i) if(xf[cellXF_bool[i]])
						xf[cellXF_bool[i]] = XParser.parsexmlbool(xf[cellXF_bool[i]]);
					if(xf.numFmtId > 0x188) {
						for(i = 0x188; i > 0x3c; --i) if(styleObj.styles.NumberFmt[xf.numFmtId] == styleObj.styles.NumberFmt[i]) { xf.numFmtId = i; break; }
					}
					styleObj.styles.CellXf.push(xf); break;
				case '</xf>': break;

				/* 18.8.1 alignment CT_CellAlignment */
				case '<alignment': case '<alignment/>':
					var alignment = {};
					if(y.vertical) alignment.vertical = y.vertical;
					if(y.horizontal) alignment.horizontal = y.horizontal;
					if(y.textRotation != null) alignment.textRotation = y.textRotation;
					if(y.indent) alignment.indent = y.indent;
					if(y.wrapText) alignment.wrapText = y.wrapText;
					xf.alignment = alignment;
					break;
				case '</alignment>': break;

				/* 18.8.33 protection CT_CellProtection */
				case '<protection': case '</protection>': case '<protection/>': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default: 
				break;
			}
		})
		XParser.xf = xf
		return styleObj.styles
	},
	
	parse_fills: function(str,styleObj) {
		var themes = styleObj.themes
		var opts = XParser.opts
		var styles = styleObj.styles
		var fill = {};
		var pass = false;
		str.match(XParser.tagregex).forEach(function(x) {
			var y = XParser.parsexmltag(x);
			switch(XParser.strip_ns(y[0])) {
				case '<fills': case '<fills>': case '</fills>': break;

				/* 18.8.20 fill CT_Fill */
				case '<fill>': case '<fill': case '<fill/>':
					fill = {}; styles.Fills.push(fill); break;
				case '</fill>': break;

				/* 18.8.24 gradientFill CT_GradientFill */
				case '<gradientFill>': break;
				case '<gradientFill':
				case '</gradientFill>': styles.Fills.push(fill); fill = {}; break;

				/* 18.8.32 patternFill CT_PatternFill */
				case '<patternFill': case '<patternFill>':
					if(y.patternType) fill.patternType = y.patternType;
					break;
				case '<patternFill/>': case '</patternFill>': break;

				/* 18.8.3 bgColor CT_Color */
				case '<bgColor':
					if(!fill.bgColor) fill.bgColor = {};
					if(y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
					if(y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
					if(y.tint) fill.bgColor.tint = parseFloat(y.tint);
					/* Excel uses ARGB strings */
					if(y.rgb) fill.bgColor.rgb = y.rgb.slice(-6);
					break;
				case '<bgColor/>': case '</bgColor>': break;

				/* 18.8.19 fgColor CT_Color */
				case '<fgColor':
					if(!fill.fgColor) fill.fgColor = {};
					if(y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
					if(y.tint) fill.fgColor.tint = parseFloat(y.tint);
					/* Excel uses ARGB strings */
					if(y.rgb) fill.fgColor.rgb = y.rgb.slice(-6);
					break;
				case '<fgColor/>': case '</fgColor>': break;

				/* 18.8.38 stop CT_GradientStop */
				case '<stop': case '<stop/>': break;
				case '</stop>': break;

				/* 18.8.? color CT_Color */
				case '<color': case '<color/>': break;
				case '</color>': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default: if(opts && opts.WTF) {
					if(!pass) throw new Error('unrecognized ' + y[0] + ' in fills');
				}
			}
		});
		return styles
	},
	
	parse_numFmts: function(str,styleObj) {
		var themes = styleObj.themes
		var opts = XParser.opts
		var styles = styleObj.styles
		//var k/*Array<number>*/ = (keys(SSF._table));
		var k = Object.keys(SSF._table);
		for(var i=0; i < k.length; ++i) styles.NumberFmt[k[i]] = SSF._table[k[i]];
		var m = str.match(XParser.tagregex);
		if(!m) return;
		for(i=0; i < m.length; ++i) {
			var y = XParser.parsexmltag(m[i]);
			switch(XParser.strip_ns(y[0])) {
				case '<numFmts': case '</numFmts>': case '<numFmts/>': case '<numFmts>': break;
				case '<numFmt': {
					var f=XParser.unescapexml(XParser.utf8read(y.formatCode)), j=parseInt(y.numFmtId,10);
					styles.NumberFmt[j] = f;
					if(j>0) {
						if(j > 0x188) {
							for(j = 0x188; j > 0x3c; --j) if(styles.NumberFmt[j] == null) break;
							styles.NumberFmt[j] = f;
						}
						SSF.load(f,j);
					}
				} break;
				case '</numFmt>': break;
				default: if(opts.WTF) throw new Error('unrecognized ' + y[0] + ' in numFmts');
			}
		}
		return styles
	},
	
	parse_fonts: function(str,styleObj) {
		var themes = styleObj.themes
		var opts = XParser.opts
		var styles = styleObj.styles
		var font = {};
		var pass = false;
		str.match(XParser.tagregex).forEach(function(x) {
			var y = XParser.parsexmltag(x);
			switch(XParser.strip_ns(y[0])) {
				case '<fonts': case '<fonts>': case '</fonts>': break;

				/* 18.8.22 font CT_Font */
				case '<font': case '<font>': break;
				case '</font>': case '<font/>':
					styles.Fonts.push(font);
					font = {};
					break;

				/* 18.8.29 name CT_FontName */
				case '<name': if(y.val) font.name = y.val; break;
				case '<name/>': case '</name>': break;

				/* 18.8.2  b CT_BooleanProperty */
				case '<b': font.bold = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<b/>': font.bold = 1; break;

				/* 18.8.26 i CT_BooleanProperty */
				case '<i': font.italic = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<i/>': font.italic = 1; break;

				/* 18.4.13 u CT_UnderlineProperty */
				case '<u':
					switch(y.val) {
						case "none": font.underline = 0x00; break;
						case "single": font.underline = 0x01; break;
						case "double": font.underline = 0x02; break;
						case "singleAccounting": font.underline = 0x21; break;
						case "doubleAccounting": font.underline = 0x22; break;
					} break;
				case '<u/>': font.underline = 1; break;

				/* 18.4.10 strike CT_BooleanProperty */
				case '<strike': font.strike = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<strike/>': font.strike = 1; break;

				/* 18.4.2  outline CT_BooleanProperty */
				case '<outline': font.outline = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<outline/>': font.outline = 1; break;

				/* 18.8.36 shadow CT_BooleanProperty */
				case '<shadow': font.shadow = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<shadow/>': font.shadow = 1; break;

				/* 18.8.12 condense CT_BooleanProperty */
				case '<condense': font.condense = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<condense/>': font.condense = 1; break;

				/* 18.8.17 extend CT_BooleanProperty */
				case '<extend': font.extend = y.val ? XParser.parsexmlbool(y.val) : 1; break;
				case '<extend/>': font.extend = 1; break;

				/* 18.4.11 sz CT_FontSize */
				case '<sz': if(y.val) font.sz = +y.val; break;
				case '<sz/>': case '</sz>': break;

				/* 18.4.14 vertAlign CT_VerticalAlignFontProperty */
				case '<vertAlign': if(y.val) font.vertAlign = y.val; break;
				case '<vertAlign/>': case '</vertAlign>': break;

				/* 18.8.18 family CT_FontFamily */
				case '<family': if(y.val) font.family = parseInt(y.val,10); break;
				case '<family/>': case '</family>': break;

				/* 18.8.35 scheme CT_FontScheme */
				case '<scheme': if(y.val) font.scheme = y.val; break;
				case '<scheme/>': case '</scheme>': break;

				/* 18.4.1 charset CT_IntProperty */
				case '<charset':
					if(y.val == '1') break;
					y.codepage = CS2CP[parseInt(y.val, 10)];
					break;

				/* 18.?.? color CT_Color */
				case '<color':
					if(!font.color) font.color = {};
					if(y.auto) font.color.auto = XParser.parsexmlbool(y.auto);

					if(y.rgb) font.color.rgb = y.rgb.slice(-6);
					else if(y.indexed) {
						font.color.index = parseInt(y.indexed, 10);
						var icv = XLSIcv[font.color.index];
						if(font.color.index == 81) icv = XLSIcv[1];
						if(!icv) throw new Error(x);
						font.color.rgb = icv[0].toString(16) + icv[1].toString(16) + icv[2].toString(16);
					} else if(y.theme) {
						font.color.theme = parseInt(y.theme, 10);
						if(y.tint) font.color.tint = parseFloat(y.tint);
						if(y.theme && themes.themeElements && themes.themeElements.clrScheme) {
							font.color.rgb = XParser.rgb_tint(themes.themeElements.clrScheme[font.color.theme].rgb, font.color.tint || 0);
						}
					}

					break;
				case '<color/>': case '</color>': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default: if(opts && opts.WTF) {
					if(!pass) throw new Error('unrecognized ' + y[0] + ' in fonts');
				}
			}
		});
		return styles
	},
	
	parse_borders: function(str,styleObj) {
		var themes = styleObj.themes
		var opts = XParser.opts
		var styles = styleObj.styles
		var border = {}/*, sub_border = {}*/;
		var pass = false;
		str.match(XParser.tagregex).forEach(function(x) {
			var y = XParser.parsexmltag(x);
			switch(XParser.strip_ns(y[0])) {
				case '<borders': case '<borders>': case '</borders>': break;

				/* 18.8.4 border CT_Border */
				case '<border': case '<border>': case '<border/>':
					border = {};
					if (y.diagonalUp) { border.diagonalUp = y.diagonalUp; }
					if (y.diagonalDown) { border.diagonalDown = y.diagonalDown; }
					styles.Borders.push(border);
					break;
				case '</border>': break;

				/* note: not in spec, appears to be CT_BorderPr */
				case '<left/>': break;
				case '<left': case '<left>': break;
				case '</left>': break;

				/* note: not in spec, appears to be CT_BorderPr */
				case '<right/>': break;
				case '<right': case '<right>': break;
				case '</right>': break;

				/* 18.8.43 top CT_BorderPr */
				case '<top/>': break;
				case '<top': case '<top>': break;
				case '</top>': break;

				/* 18.8.6 bottom CT_BorderPr */
				case '<bottom/>': break;
				case '<bottom': case '<bottom>': break;
				case '</bottom>': break;

				/* 18.8.13 diagonal CT_BorderPr */
				case '<diagonal': case '<diagonal>': case '<diagonal/>': break;
				case '</diagonal>': break;

				/* 18.8.25 horizontal CT_BorderPr */
				case '<horizontal': case '<horizontal>': case '<horizontal/>': break;
				case '</horizontal>': break;

				/* 18.8.44 vertical CT_BorderPr */
				case '<vertical': case '<vertical>': case '<vertical/>': break;
				case '</vertical>': break;

				/* 18.8.37 start CT_BorderPr */
				case '<start': case '<start>': case '<start/>': break;
				case '</start>': break;

				/* 18.8.16 end CT_BorderPr */
				case '<end': case '<end>': case '<end/>': break;
				case '</end>': break;

				/* 18.8.? color CT_Color */
				case '<color': case '<color>': break;
				case '<color/>': case '</color>': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default: if(opts && opts.WTF) {
					if(!pass) throw new Error('unrecognized ' + y[0] + ' in borders');
				}
			}
		});
		return styles
	},

	
	parsexmltag: function(tag, skip_root) {
		var attregexg = /([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
		var z = ({});
		var eq = 0, c = 0;
		for(; eq < tag.length; eq++) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
		if(!skip_root) z[0] = tag.slice(0, eq);
		if(eq === tag.length) return z;
		var m = tag.match(attregexg), j=0, v="", i=0, q="", cc="", quot = 1;
		if(m){ 
			var i = 0;
			for(; i < m.length; i++) {
				cc = m[i];
				var c = 0;
				for(; c < cc.length; c++)
					if(cc.charCodeAt(c) === 61) break;
				q = cc.slice(0,c).trim();
				while(cc.charCodeAt(c+1) == 32) 
					++c;
					quot = ((eq=cc.charCodeAt(c+1)) == 34 || eq == 39) ? 1 : 0;
					v = cc.slice(c+1+quot, cc.length-quot);
				var j = 0;
				for(;j < q.length;j++) 
					if(q.charCodeAt(j) === 58) break;
				if(j===q.length) {
					if(q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_")); // from ods
					z[q] = v;
					z[q.toLowerCase()] = v;
				}
				else {
					var k = (j===5 && q.slice(0,5)==="xmlns"?"xmlns":"")+q.slice(j+1);
					if(z[k] && q.slice(j-3,j) == "ext") continue; // from ods
					z[k] = v;
					z[k.toLowerCase()] = v;
				}
			}
		}
		return z;
	},	
	
	parse_si: function(x, opts) {
		var html = opts ? opts.cellHTML : true;
		var z = {};
		if(!x) return null;
		//var y;
		/* 18.4.12 t ST_Xstring (Plaintext String) */
		// TODO: is whitespace actually valid here?
		if(x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
			z.t = XParser.unescapexml(XParser.utf8read(x.slice(x.indexOf(">")+1).split(/<\/(?:\w+:)?t>/)[0]||""));
			z.r = XParser.utf8read(x);
			if(html) z.h = XParser.escapehtml(z.t);
		}
		/* 18.4.4 r CT_RElt (Rich Text Run) */
		else if((/*y = */x.match(XParser.sirregex))) {
			z.r = XParser.utf8read(x);
			z.t = XParser.unescapexml(XParser.utf8read((x.replace(XParser.sirphregex, '').match(XParser.sitregex)||[]).join("").replace(tagregex,"")));
			if(html) z.h = parse_rs(z.r);
		}
		/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
		/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
		return z;
	},
				
	parse_rpr: function(rpr, intro, outro) {
		var font = {}, cp = 65001, align = "";
		var pass = false;
		var m = rpr.match(tagregex), i = 0;
		if(m) for(;i!=m.length; ++i) {
			var y = XParser.parsexmltag(m[i]);
			switch(y[0].replace(/\w*:/g,"")) {
				/* 18.8.12 condense CT_BooleanProperty */
				/* ** not required . */
				case '<condense': break;
				/* 18.8.17 extend CT_BooleanProperty */
				/* ** not required . */
				case '<extend': break;
				/* 18.8.36 shadow CT_BooleanProperty */
				/* ** not required . */
				case '<shadow':
					if(!y.val) break;
					/* falls through */
				case '<shadow>':
				case '<shadow/>': font.shadow = 1; break;
				case '</shadow>': break;

				/* 18.4.1 charset CT_IntProperty TODO */
				case '<charset':
					if(y.val == '1') break;
					cp = CS2CP[parseInt(y.val, 10)];
					break;

				/* 18.4.2 outline CT_BooleanProperty TODO */
				case '<outline':
					if(!y.val) break;
					/* falls through */
				case '<outline>':
				case '<outline/>': font.outline = 1; break;
				case '</outline>': break;

				/* 18.4.5 rFont CT_FontName */
				case '<rFont': font.name = y.val; break;

				/* 18.4.11 sz CT_FontSize */
				case '<sz': font.sz = y.val; break;

				/* 18.4.10 strike CT_BooleanProperty */
				case '<strike':
					if(!y.val) break;
					/* falls through */
				case '<strike>':
				case '<strike/>': font.strike = 1; break;
				case '</strike>': break;

				/* 18.4.13 u CT_UnderlineProperty */
				case '<u':
					if(!y.val) break;
					switch(y.val) {
						case 'double': font.uval = "double"; break;
						case 'singleAccounting': font.uval = "single-accounting"; break;
						case 'doubleAccounting': font.uval = "double-accounting"; break;
					}
					/* falls through */
				case '<u>':
				case '<u/>': font.u = 1; break;
				case '</u>': break;

				/* 18.8.2 b */
				case '<b':
					if(y.val == '0') break;
					/* falls through */
				case '<b>':
				case '<b/>': font.b = 1; break;
				case '</b>': break;

				/* 18.8.26 i */
				case '<i':
					if(y.val == '0') break;
					/* falls through */
				case '<i>':
				case '<i/>': font.i = 1; break;
				case '</i>': break;

				/* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
				case '<color':
					if(y.rgb) font.color = y.rgb.slice(2,8);
					break;

				/* 18.8.18 family ST_FontFamily */
				case '<family': font.family = y.val; break;

				/* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
				case '<vertAlign': align = y.val; break;

				/* 18.8.35 scheme CT_FontScheme TODO */
				case '<scheme': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default:
					if(y[0].charCodeAt(1) !== 47 && !pass) throw new Error('Unrecognized rich format ' + y[0]);
			}
		}
		var style = [];

		if(font.u) style.push("text-decoration: underline;");
		if(font.uval) style.push("text-underline-style:" + font.uval + ";");
		if(font.sz) style.push("font-size:" + font.sz + "pt;");
		if(font.outline) style.push("text-effect: outline;");
		if(font.shadow) style.push("text-shadow: auto;");
		intro.push('<span style="' + style.join("") + '">');

		if(font.b) { intro.push("<b>"); outro.push("</b>"); }
		if(font.i) { intro.push("<i>"); outro.push("</i>"); }
		if(font.strike) { intro.push("<s>"); outro.push("</s>"); }

		if(align == "superscript") align = "sup";
		else if(align == "subscript") align = "sub";
		if(align != "") { intro.push("<" + align + ">"); outro.push("</" + align + ">"); }

		outro.push("</span>");
		return cp;
	},
	
	parse_r: function(r) {
		var rpregex = matchtag("rPr")
		var tregex = matchtag("t")
		var terms = [[],"",[]];
		/* 18.4.12 t ST_Xstring */
		var t = r.match(tregex)/*, cp = 65001*/;
		if(!t) return "";
		terms[1] = t[1];

		var rpr = r.match(rpregex);
		if(rpr) /*cp = */parse_rpr(rpr[1], terms[0], terms[2]);

		return terms[0].join("") + terms[1].replace(XParser.nlregex,'<br/>') + terms[2].join("");
	},			
	parse_rs: function(rs) {
		return rs.replace(XParser.rregex,"").split(XParser.rend).map(parse_r).join("");
	},
	
	strip_ns: function(x) { 
		var nsregex2 = /<(\/?)\w+:/
		return x.replace(XParser.nsregex2, "<$1"); 
	},
	parsexmlbool: function(value) {
		switch(value) {
			case 1: case true: case '1': case 'true': case 'TRUE': return true;
			/* case '0': case 'false': case 'FALSE':*/
			default: return false;
		}
	},

/* parses a date as a local date */
	parseDate: function(str, fixdate) {
		var d = new Date(str);
		if(good_pd) {
			if(fixdate > 0) d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
			else if(fixdate < 0) d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
			return d;
		}
		if(str instanceof Date) return str;
		if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
			var s = d.getFullYear();
			if(str.indexOf("" + s) > -1) return d;
			d.setFullYear(d.getFullYear() + 100); return d;
		}
		var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
		var out = new Date(+n[0], +n[1] - 1, +n[2], (+n[3]||0), (+n[4]||0), (+n[5]||0));
		if(str.indexOf("Z") > -1) out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
		return out;
	},
	
	datenum: function(v, date1904) {
		var epoch = v.getTime();
		if(date1904) epoch -= 1462*24*60*60*1000;
		return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	},
	
	//honestly have no idea what this is doing... but xlsx had it
	px2pt: function(px) { 
		return px * 96 / 96; 
	},
	pt2px: function(pt) { 
		return pt * 96 / 96; 
	},
	encode_cell: function(cell) { 
		return XParser.encode_col(cell.c) + XParser.encode_row(cell.r); 
	},
	encode_col: function(col) { 
		var s=""; 
		for(++col; col; col=Math.floor((col-1)/26))
			s = String.fromCharCode(((col-1)%26) + 65) + s; 
		return s; 
	},
	encode_row: function(row) { 
		return "" + (row + 1); 
	},
	decode_cell: function(cstr) { 
		var splt = XParser.split_cell(cstr); 
		return { c: XParser.decode_col(splt[0]), r: XParser.decode_row(splt[1]) }; 
	},
	split_cell: function(cstr) { 
		return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); 
	},
	decode_row: function(rowstr) { 
		return parseInt(XParser.unfix_row(rowstr),10) - 1; 
	},
	decode_range: function(range) { 
		var x = range.split(":").map(XParser.decode_cell);
		return {s:x[0],e:x[x.length-1]}; 
	},
	unfix_row: function(cstr) { 
		return cstr.replace(/\$(\d+)$/,"$1"); 
	},
	unfix_col: function(cstr) { 
		return cstr.replace(/^\$([A-Z])/,"$1"); 
	},
	decode_col: function(colstr) { 
		var c = XParser.unfix_col(colstr), d = 0, i = 0; 
		for(; i !== c.length; ++i) 
			d = 26*d + c.charCodeAt(i) - 64; 
		return d - 1; 
	},
	numdate: function(v) {
		var out = new Date();
		out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
		return out;
	},
	rgb_tint: function(hex, tint) {
		if(tint === 0) return hex;
		var hsl = XParser.rgb2HSL(XParser.hex2RGB(hex));
		if (tint < 0) hsl[2] = hsl[2] * (1 + tint);
		else hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
		return XParser.rgb2Hex(XParser.hsl2RGB(hsl));
	},
	rgb2HSL: function(rgb) {
		var R = rgb[0]/255, G = rgb[1]/255, B=rgb[2]/255;
		var M = Math.max(R, G, B), m = Math.min(R, G, B), C = M - m;
		if(C === 0) return [0, 0, R];

		var H6 = 0, S = 0, L2 = (M + m);
		S = C / (L2 > 1 ? 2 - L2 : L2);
		switch(M){
			case R: H6 = ((G - B) / C + 6)%6; break;
			case G: H6 = ((B - R) / C + 2); break;
			case B: H6 = ((R - G) / C + 4); break;
		}
		return [H6 / 6, S, L2 / 2];
	},
	hex2RGB: function(h) {
		var o = h.slice(h[0]==="#"?1:0).slice(0,6);
		return [parseInt(o.slice(0,2),16),parseInt(o.slice(2,4),16),parseInt(o.slice(4,6),16)];
	},
	rgb2Hex: function(rgb) {
		for(var i=0,o=1; i!=3; ++i) o = o*256 + (rgb[i]>255?255:rgb[i]<0?0:rgb[i]);
		return o.toString(16).toUpperCase().slice(1);
	},
	hsl2RGB: function(hsl){
		var H = hsl[0], S = hsl[1], L = hsl[2];
		var C = S * 2 * (L < 0.5 ? L : 1 - L), m = L - C/2;
		var rgb = [m,m,m], h6 = 6*H;

		var X;
		if(S !== 0) switch(h6|0) {
			case 0: case 6: X = C * h6; rgb[0] += C; rgb[1] += X; break;
			case 1: X = C * (2 - h6);   rgb[0] += X; rgb[1] += C; break;
			case 2: X = C * (h6 - 2);   rgb[1] += C; rgb[2] += X; break;
			case 3: X = C * (4 - h6);   rgb[1] += X; rgb[2] += C; break;
			case 4: X = C * (h6 - 4);   rgb[2] += C; rgb[0] += X; break;
			case 5: X = C * (6 - h6);   rgb[2] += X; rgb[0] += C; break;
		}
		for(var i = 0; i != 3; ++i) rgb[i] = Math.round(rgb[i]*255);
		return rgb;
	},

	unescapexml:function(text) {
		
		var s = text + '', i = s.indexOf("<![CDATA[");
		if(i == -1) 
			return s.replace(XParser.encregex, function($$, $1) { 
			return XParser.encodings[$$]||String.fromCharCode(parseInt($1,$$.indexOf("x")>-1?16:10))||$$; }).replace(XParser.coderegex,function(m,c) {
				return String.fromCharCode(parseInt(c,16));
			});
		var j = s.indexOf("]]>");
		return XParser.unescapexml(s.slice(0, i)) + s.slice(i+9,j) + XParser.unescapexml(s.slice(j+3));
	},
	escapehtml: function(text){
		var s = text + '';
		return s.replace(XParser.decregex, function(y) { return XParser.evert(XParser.encodings)[y]; }).replace(/\n/g, "<br/>").replace(XParser.htmlcharegex,function(s) { return "&#x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + ";"; });
	},
	evert: function(obj) {
		var o = ([]), K = Object.keys(obj);
		for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
		return o;
	},
	utf8read: function(orig) {
		var out = "", i = 0, c = 0, d = 0, e = 0, f = 0, w = 0;
		while (i < orig.length) {
			c = orig.charCodeAt(i++);
			if (c < 128) { out += String.fromCharCode(c); continue; }
			d = orig.charCodeAt(i++);
			if (c>191 && c<224) { f = ((c & 31) << 6); f |= (d & 63); out += String.fromCharCode(f); continue; }
			e = orig.charCodeAt(i++);
			if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
			f = orig.charCodeAt(i++);
			w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63))-65536;
			out += String.fromCharCode(0xD800 + ((w>>>10)&1023));
			out += String.fromCharCode(0xDC00 + (w&1023));
		}
		return out;
	},
	
	safe_decode_range: function(range) {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		var idx = 0, i = 0, cc = 0;
		var len = range.length;
		for(idx = 0; i < len; ++i) {
			if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
			idx = 26*idx + cc;
		}
		o.s.c = --idx;

		for(idx = 0; i < len; ++i) {
			if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
			idx = 10*idx + cc;
		}
		o.s.r = --idx;

		if(i === len || range.charCodeAt(++i) === 58) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }

		for(idx = 0; i != len; ++i) {
			if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
			idx = 26*idx + cc;
		}
		o.e.c = --idx;

		for(idx = 0; i != len; ++i) {
			if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
			idx = 10*idx + cc;
		}
		o.e.r = --idx;
		return o;
	},
	
	_xlfn: function(f) {
		return f.replace(/_xlfn\./g,"");
	},
	
	shift_formula_str: function(f, delta) {
		return f.replace(XParser.crefregex, function($0, $1, $2, $3, $4, $5) {
			return $1+($2=="$" ? $2+$3 : XParser.encode_col(XParser.decode_col($3)+delta.c))+($4=="$" ? $4+$5 : XParser.encode_row(XParser.decode_row($5) + delta.r));
		});
	},

	shift_formula_xlsx: function(f, range, cell) {
		var r = XParser.decode_range(range), s = r.s, c = XParser.decode_cell(cell);
		var delta = {r:c.r - s.r, c:c.c - s.c};
		return XParser.shift_formula_str(f, delta);
	},
	
	/** CREATING THE THEMES  **/
	parse_clrScheme: function(t, themes, opts) {
		themes.themeElements.clrScheme = [];
		var color = {};
		(t[0].match(XParser.tagregex)||[]).forEach(function(x) {
			var y = XParser.parsexmltag(x);
			switch(y[0]) {
				/* 20.1.6.2 clrScheme (Color Scheme) CT_ColorScheme */
				case '<a:clrScheme': case '</a:clrScheme>': break;

				/* 20.1.2.3.32 srgbClr CT_SRgbColor */
				case '<a:srgbClr':
					color.rgb = y.val; break;

				/* 20.1.2.3.33 sysClr CT_SystemColor */
				case '<a:sysClr':
					color.rgb = y.lastClr; break;

				/* 20.1.4.1.1 accent1 (Accent 1) */
				/* 20.1.4.1.2 accent2 (Accent 2) */
				/* 20.1.4.1.3 accent3 (Accent 3) */
				/* 20.1.4.1.4 accent4 (Accent 4) */
				/* 20.1.4.1.5 accent5 (Accent 5) */
				/* 20.1.4.1.6 accent6 (Accent 6) */
				/* 20.1.4.1.9 dk1 (Dark 1) */
				/* 20.1.4.1.10 dk2 (Dark 2) */
				/* 20.1.4.1.15 folHlink (Followed Hyperlink) */
				/* 20.1.4.1.19 hlink (Hyperlink) */
				/* 20.1.4.1.22 lt1 (Light 1) */
				/* 20.1.4.1.23 lt2 (Light 2) */
				case '<a:dk1>': case '</a:dk1>':
				case '<a:lt1>': case '</a:lt1>':
				case '<a:dk2>': case '</a:dk2>':
				case '<a:lt2>': case '</a:lt2>':
				case '<a:accent1>': case '</a:accent1>':
				case '<a:accent2>': case '</a:accent2>':
				case '<a:accent3>': case '</a:accent3>':
				case '<a:accent4>': case '</a:accent4>':
				case '<a:accent5>': case '</a:accent5>':
				case '<a:accent6>': case '</a:accent6>':
				case '<a:hlink>': case '</a:hlink>':
				case '<a:folHlink>': case '</a:folHlink>':
					if (y[0].charAt(1) === '/') {
						themes.themeElements.clrScheme.push(color);
						color = {};
					} else {
						color.name = y[0].slice(3, y[0].length - 1);
					}
					break;

				default: if(opts && opts.WTF) throw new Error('Unrecognized ' + y[0] + ' in clrScheme');
			}
		});
	},
	
	parse_themeElements: function(data, themes, opts) {
		themes.themeElements = {};
		var clrsregex = /<a:clrScheme([^>]*)>[\s\S]*<\/a:clrScheme>/;
		var fntsregex = /<a:fontScheme([^>]*)>[\s\S]*<\/a:fontScheme>/;
		var fmtsregex = /<a:fmtScheme([^>]*)>[\s\S]*<\/a:fmtScheme>/;
		var t;

		[
			/* clrScheme CT_ColorScheme */
			['clrScheme', clrsregex, XParser.parse_clrScheme],
			/* fontScheme CT_FontScheme */
			['fontScheme', fntsregex, function parse_fontScheme() { }],
			/* fmtScheme CT_StyleMatrix */
			['fmtScheme', fmtsregex, function parse_fmtScheme() { }]
		].forEach(function(m) {
			if(!(t=data.match(m[1]))) throw new Error(m[0] + ' not found in themeElements');
			m[2](t, themes, opts);
		});
	},

	/* 14.2.7 Theme Part */
	parse_theme_xml: function(data, opts) {
		var themeltregex = /<a:themeElements([^>]*)>[\s\S]*<\/a:themeElements>/;
		/* 20.1.6.9 theme CT_OfficeStyleSheet */
		if(!data || data.length === 0) return XParser.parse_theme_xml(XParser.write_theme());
		var t;
		var themes = {};

		/* themeElements CT_BaseStyles */
		if(!(t=data.match(themeltregex))) throw new Error('themeElements not found in theme');
		XParser.parse_themeElements(t[0], themes, opts);

		return themes;
	},
	write_theme: function(Themes, opts) {
		if(opts && opts.themeXLSX) return opts.themeXLSX;
		var o = [XML_HEADER];
		o[o.length] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
		o[o.length] =  '<a:themeElements>';

		o[o.length] =   '<a:clrScheme name="Office">';
		o[o.length] =    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
		o[o.length] =    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
		o[o.length] =    '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
		o[o.length] =    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
		o[o.length] =    '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
		o[o.length] =    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
		o[o.length] =    '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
		o[o.length] =    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
		o[o.length] =    '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
		o[o.length] =    '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
		o[o.length] =    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
		o[o.length] =    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
		o[o.length] =   '</a:clrScheme>';

		o[o.length] =   '<a:fontScheme name="Office">';
		o[o.length] =    '<a:majorFont>';
		o[o.length] =     '<a:latin typeface="Cambria"/>';
		o[o.length] =     '<a:ea typeface=""/>';
		o[o.length] =     '<a:cs typeface=""/>';
		o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
		o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
		o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
		o[o.length] =     '<a:font script="Arab" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Hebr" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
		o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
		o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
		o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
		o[o.length] =     '<a:font script="Khmr" typeface="MoolBoran"/>';
		o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
		o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
		o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
		o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
		o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
		o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
		o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
		o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
		o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
		o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
		o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		o[o.length] =     '<a:font script="Viet" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
		o[o.length] =    '</a:majorFont>';
		o[o.length] =    '<a:minorFont>';
		o[o.length] =     '<a:latin typeface="Calibri"/>';
		o[o.length] =     '<a:ea typeface=""/>';
		o[o.length] =     '<a:cs typeface=""/>';
		o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
		o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
		o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
		o[o.length] =     '<a:font script="Arab" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Hebr" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
		o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
		o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
		o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
		o[o.length] =     '<a:font script="Khmr" typeface="DaunPenh"/>';
		o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
		o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
		o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
		o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
		o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
		o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
		o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
		o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
		o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
		o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
		o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		o[o.length] =     '<a:font script="Viet" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
		o[o.length] =    '</a:minorFont>';
		o[o.length] =   '</a:fontScheme>';

		o[o.length] =   '<a:fmtScheme name="Office">';
		o[o.length] =    '<a:fillStyleLst>';
		o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:lin ang="16200000" scaled="1"/>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:lin ang="16200000" scaled="0"/>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =    '</a:fillStyleLst>';
		o[o.length] =    '<a:lnStyleLst>';
		o[o.length] =     '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =     '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =     '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =    '</a:lnStyleLst>';
		o[o.length] =    '<a:effectStyleLst>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =      '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
		o[o.length] =      '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =    '</a:effectStyleLst>';
		o[o.length] =    '<a:bgFillStyleLst>';
		o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =    '</a:bgFillStyleLst>';
		o[o.length] =   '</a:fmtScheme>';
		o[o.length] =  '</a:themeElements>';

		o[o.length] =  '<a:objectDefaults>';
		o[o.length] =   '<a:spDef>';
		o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
		o[o.length] =   '</a:spDef>';
		o[o.length] =   '<a:lnDef>';
		o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
		o[o.length] =   '</a:lnDef>';
		o[o.length] =  '</a:objectDefaults>';
		o[o.length] =  '<a:extraClrSchemeLst/>';
		o[o.length] = '</a:theme>';
		return o.join("");
	}
	/** END CREATING THE THEMES  **/
} //xparser


function rgbify(arr) { return arr.map(function(x) { return [(x>>16)&255,(x>>8)&255,x&255]; }); }
var XLSIcv = rgbify([
	/* Color Constants */
	0x000000,
	0xFFFFFF,
	0xFF0000,
	0x00FF00,
	0x0000FF,
	0xFFFF00,
	0xFF00FF,
	0x00FFFF,

	/* Overridable Defaults */
	0x000000,
	0xFFFFFF,
	0xFF0000,
	0x00FF00,
	0x0000FF,
	0xFFFF00,
	0xFF00FF,
	0x00FFFF,

	0x800000,
	0x008000,
	0x000080,
	0x808000,
	0x800080,
	0x008080,
	0xC0C0C0,
	0x808080,
	0x9999FF,
	0x993366,
	0xFFFFCC,
	0xCCFFFF,
	0x660066,
	0xFF8080,
	0x0066CC,
	0xCCCCFF,

	0x000080,
	0xFF00FF,
	0xFFFF00,
	0x00FFFF,
	0x800080,
	0x800000,
	0x008080,
	0x0000FF,
	0x00CCFF,
	0xCCFFFF,
	0xCCFFCC,
	0xFFFF99,
	0x99CCFF,
	0xFF99CC,
	0xCC99FF,
	0xFFCC99,

	0x3366FF,
	0x33CCCC,
	0x99CC00,
	0xFFCC00,
	0xFF9900,
	0xFF6600,
	0x666699,
	0x969696,
	0x003366,
	0x339966,
	0x003300,
	0x333300,
	0x993300,
	0x993366,
	0x333399,
	0x333333,

	/* Other entries to appease BIFF8/12 */
	0xFFFFFF, /* 0x40 icvForeground ?? */
	0x000000, /* 0x41 icvBackground ?? */
	0x000000, /* 0x42 icvFrame ?? */
	0x000000, /* 0x43 icv3D ?? */
	0x000000, /* 0x44 icv3DText ?? */
	0x000000, /* 0x45 icv3DHilite ?? */
	0x000000, /* 0x46 icv3DShadow ?? */
	0x000000, /* 0x47 icvHilite ?? */
	0x000000, /* 0x48 icvCtlText ?? */
	0x000000, /* 0x49 icvCtlScrl ?? */
	0x000000, /* 0x4A icvCtlInv ?? */
	0x000000, /* 0x4B icvCtlBody ?? */
	0x000000, /* 0x4C icvCtlFrame ?? */
	0x000000, /* 0x4D icvCtlFore ?? */
	0x000000, /* 0x4E icvCtlBack ?? */
	0x000000, /* 0x4F icvCtlNeutral */
	0x000000, /* 0x50 icvInfoBk ?? */
	0x000000 /* 0x51 icvInfoText ?? */
]);


/** XLSX had this for the date parser. not sure whats up with the hardcoding **/
var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2/19/17');
var good_pd = good_pd_date.getFullYear() == 2017;
var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;

var matchtag = (function() {
	var mtcache = ({});
	return function matchtag(f,g) {
		var t = f+"|"+(g||"");
		if(mtcache[t]) return mtcache[t];
		return (mtcache[t] = new RegExp('<(?:\\w+:)?'+f+'(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?'+f+'>',((g||""))));
	};
})();


onmessage = function(e){
	if(e.origin.indexOf('chrome-extension') == -1)
		XParser.parse_xml_from_excel(e.data.data,e.data.zip)
}




var CS2CP = ({
0:    1252, /* ANSI */
1:   65001, /* DEFAULT */
2:   65001, /* SYMBOL */
77:  10000, /* MAC */
128:   932, /* SHIFTJIS */
129:   949, /* HANGUL */
130:  1361, /* JOHAB */
134:   936, /* GB2312 */
136:   950, /* CHINESEBIG5 */
161:  1253, /* GREEK */
162:  1254, /* TURKISH */
163:  1258, /* VIETNAMESE */
177:  1255, /* HEBREW */
178:  1256, /* ARABIC */
186:  1257, /* BALTIC */
204:  1251, /* RUSSIAN */
222:   874, /* THAI */
238:  1250, /* EASTEUROPE */
255:  1252, /* OEM */
69:   6969  /* MISC */
});









/*** SSF   ***/

/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint -W041 */
var SSF = ({});
var make_ssf = function make_ssf(SSF){
SSF.version = '0.11.0';
function _strrev(x) { var o = "", i = x.length-1; while(i>=0) o += x.charAt(i--); return o; }
function fill(c,l) { var o = ""; while(o.length < l) o+=c; return o; }
function pad0(v,d){var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
function pad_(v,d){var t=""+v;return t.length>=d?t:fill(' ',d-t.length)+t;}
function rpad_(v,d){var t=""+v; return t.length>=d?t:t+fill(' ',d-t.length);}
function pad0r1(v,d){var t=""+Math.round(v); return t.length>=d?t:fill('0',d-t.length)+t;}
function pad0r2(v,d){var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
var p2_32 = Math.pow(2,32);
function pad0r(v,d){if(v>p2_32||v<-p2_32) return pad0r1(v,d); var i = Math.round(v); return pad0r2(i,d); }
function isgeneral(s, i) { i = i || 0; return s.length >= 7 + i && (s.charCodeAt(i)|32) === 103 && (s.charCodeAt(i+1)|32) === 101 && (s.charCodeAt(i+2)|32) === 110 && (s.charCodeAt(i+3)|32) === 101 && (s.charCodeAt(i+4)|32) === 114 && (s.charCodeAt(i+5)|32) === 97 && (s.charCodeAt(i+6)|32) === 108; }
var days = [
	['Sun', 'Sunday'],
	['Mon', 'Monday'],
	['Tue', 'Tuesday'],
	['Wed', 'Wednesday'],
	['Thu', 'Thursday'],
	['Fri', 'Friday'],
	['Sat', 'Saturday']
];
var months = [
	['J', 'Jan', 'January'],
	['F', 'Feb', 'February'],
	['M', 'Mar', 'March'],
	['A', 'Apr', 'April'],
	['M', 'May', 'May'],
	['J', 'Jun', 'June'],
	['J', 'Jul', 'July'],
	['A', 'Aug', 'August'],
	['S', 'Sep', 'September'],
	['O', 'Oct', 'October'],
	['N', 'Nov', 'November'],
	['D', 'Dec', 'December']
];
function init_table(t) {
	t[0]=  'General';
	t[1]=  '0';
	t[2]=  '0.00';
	t[3]=  '#,##0';
	t[4]=  '#,##0.00';
	t[9]=  '0%';
	t[10]= '0.00%';
	t[11]= '0.00E+00';
	t[12]= '# ?/?';
	t[13]= '# ??/??';
	//t[14]= 'm/d/yy';
	t[14]= 'mm/dd/yyyy';
	t[15]= 'd-mmm-yy';
	t[16]= 'd-mmm';
	t[17]= 'mmm-yy';
	t[18]= 'h:mm AM/PM';
	t[19]= 'h:mm:ss AM/PM';
	t[20]= 'h:mm';
	t[21]= 'h:mm:ss';
	t[22]= 'm/d/yy h:mm';
	t[37]= '#,##0 ;(#,##0)';
	t[38]= '#,##0 ;[Red](#,##0)';
	t[39]= '#,##0.00;(#,##0.00)';
	t[40]= '#,##0.00;[Red](#,##0.00)';
	t[45]= 'mm:ss';
	t[46]= '[h]:mm:ss';
	t[47]= 'mmss.0';
	t[48]= '##0.0E+0';
	t[49]= '@';
	t[56]= '"上午/下午 "hh"時"mm"分"ss"秒 "';
	t[65535]= 'General';
}

var table_fmt = {};
init_table(table_fmt);
function frac(x, D, mixed) {
	var sgn = x < 0 ? -1 : 1;
	var B = x * sgn;
	var P_2 = 0, P_1 = 1, P = 0;
	var Q_2 = 1, Q_1 = 0, Q = 0;
	var A = Math.floor(B);
	while(Q_1 < D) {
		A = Math.floor(B);
		P = A * P_1 + P_2;
		Q = A * Q_1 + Q_2;
		if((B - A) < 0.00000005) break;
		B = 1 / (B - A);
		P_2 = P_1; P_1 = P;
		Q_2 = Q_1; Q_1 = Q;
	}
	if(Q > D) { if(Q_1 > D) { Q = Q_2; P = P_2; } else { Q = Q_1; P = P_1; } }
	if(!mixed) return [0, sgn * P, Q];
	var q = Math.floor(sgn * P/Q);
	return [q, sgn*P - q*Q, Q];
}
function parse_date_code(v,opts,b2) {
	if(v > 2958465 || v < 0) return null;
	var date = (v|0), time = Math.floor(86400 * (v - date)), dow=0;
	var dout=[];
	var out={D:date, T:time, u:86400*(v-date)-time,y:0,m:0,d:0,H:0,M:0,S:0,q:0};
	if(Math.abs(out.u) < 1e-6) out.u = 0;
	if(opts && opts.date1904) date += 1462;
	if(out.u > 0.9999) {
		out.u = 0;
		if(++time == 86400) { out.T = time = 0; ++date; ++out.D; }
	}
	if(date === 60) {dout = b2 ? [1317,10,29] : [1900,2,29]; dow=3;}
	else if(date === 0) {dout = b2 ? [1317,8,29] : [1900,1,0]; dow=6;}
	else {
		if(date > 60) --date;
		/* 1 = Jan 1 1900 in Gregorian */
		var d = new Date(1900, 0, 1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
		dow = d.getDay();
		if(date < 60) dow = (dow + 6) % 7;
		if(b2) dow = fix_hijri(d, dout);
	}
	out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
	out.S = time % 60; time = Math.floor(time / 60);
	out.M = time % 60; time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
	return out;
}
SSF.parse_date_code = parse_date_code;
var basedate = new Date(1899, 11, 31, 0, 0, 0);
var dnthresh = basedate.getTime();
var base1904 = new Date(1900, 2, 1, 0, 0, 0);
function datenum_local(v, date1904) {
	var epoch = v.getTime();
	if(date1904) epoch -= 1461*24*60*60*1000;
	else if(v >= base1904) epoch += 24*60*60*1000;
	return (epoch - (dnthresh + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000)) / (24 * 60 * 60 * 1000);
}
function general_fmt_int(v) { return v.toString(10); }
SSF._general_int = general_fmt_int;
var general_fmt_num = (function make_general_fmt_num() {
var gnr1 = /\.(\d*[1-9])0+$/, gnr2 = /\.0*$/, gnr4 = /\.(\d*[1-9])0+/, gnr5 = /\.0*[Ee]/, gnr6 = /(E[+-])(\d)$/;
function gfn2(v) {
	var w = (v<0?12:11);
	var o = gfn5(v.toFixed(12)); if(o.length <= w) return o;
	o = v.toPrecision(10); if(o.length <= w) return o;
	return v.toExponential(5);
}
function gfn3(v) {
	var o = v.toFixed(11).replace(gnr1,".$1");
	if(o.length > (v<0?12:11)) o = v.toPrecision(6);
	return o;
}
function gfn4(o) {
	for(var i = 0; i != o.length; ++i) if((o.charCodeAt(i) | 0x20) === 101) return o.replace(gnr4,".$1").replace(gnr5,"E").replace("e","E").replace(gnr6,"$10$2");
	return o;
}
function gfn5(o) {
	return o.indexOf(".") > -1 ? o.replace(gnr2,"").replace(gnr1,".$1") : o;
}
return function general_fmt_num(v) {
	var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;
	if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
	else if(Math.abs(V) <= 9) o = gfn2(v);
	else if(V === 10) o = v.toFixed(10).substr(0,12);
	else o = gfn3(v);
	return gfn5(gfn4(o));
};})();
SSF._general_num = general_fmt_num;
function general_fmt(v, opts) {
	switch(typeof v) {
		case 'string': return v;
		case 'boolean': return v ? "TRUE" : "FALSE";
		case 'number': return (v|0) === v ? v.toString(10) : general_fmt_num(v);
		case 'undefined': return "";
		case 'object':
			if(v == null) return "";
			if(v instanceof Date) return format(14, datenum_local(v, opts && opts.date1904), opts);
	}
	throw new Error("unsupported value in General format: " + v);
}
SSF._general = general_fmt;
function fix_hijri() { return 0; }
/*jshint -W086 */
function write_date(type, fmt, val, ss0) {
	var o="", ss=0, tt=0, y = val.y, out, outl = 0;
	switch(type) {
		case 98: /* 'b' buddhist year */
			y = val.y + 543;
			/* falls through */
		case 121: /* 'y' year */
		switch(fmt.length) {
			case 1: case 2: out = y % 100; outl = 2; break;
			default: out = y % 10000; outl = 4; break;
		} break;
		case 109: /* 'm' month */
		switch(fmt.length) {
			case 1: case 2: out = val.m; outl = fmt.length; break;
			case 3: return months[val.m-1][1];
			case 5: return months[val.m-1][0];
			default: return months[val.m-1][2];
		} break;
		case 100: /* 'd' day */
		switch(fmt.length) {
			case 1: case 2: out = val.d; outl = fmt.length; break;
			case 3: return days[val.q][0];
			default: return days[val.q][1];
		} break;
		case 104: /* 'h' 12-hour */
		switch(fmt.length) {
			case 1: case 2: out = 1+(val.H+11)%12; outl = fmt.length; break;
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 72: /* 'H' 24-hour */
		switch(fmt.length) {
			case 1: case 2: out = val.H; outl = fmt.length; break;
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 77: /* 'M' minutes */
		switch(fmt.length) {
			case 1: case 2: out = val.M; outl = fmt.length; break;
			default: throw 'bad minute format: ' + fmt;
		} break;
		case 115: /* 's' seconds */
			if(fmt != 's' && fmt != 'ss' && fmt != '.0' && fmt != '.00' && fmt != '.000') throw 'bad second format: ' + fmt;
			if(val.u === 0 && (fmt == "s" || fmt == "ss")) return pad0(val.S, fmt.length);
if(ss0 >= 2) tt = ss0 === 3 ? 1000 : 100;
			else tt = ss0 === 1 ? 10 : 1;
			ss = Math.round((tt)*(val.S + val.u));
			if(ss >= 60*tt) ss = 0;
			if(fmt === 's') return ss === 0 ? "0" : ""+ss/tt;
			o = pad0(ss,2 + ss0);
			if(fmt === 'ss') return o.substr(0,2);
			return "." + o.substr(2,fmt.length-1);
		case 90: /* 'Z' absolute time */
		switch(fmt) {
			case '[h]': case '[hh]': out = val.D*24+val.H; break;
			case '[m]': case '[mm]': out = (val.D*24+val.H)*60+val.M; break;
			case '[s]': case '[ss]': out = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
			default: throw 'bad abstime format: ' + fmt;
		} outl = fmt.length === 3 ? 1 : 2; break;
		case 101: /* 'e' era */
			out = y; outl = 1;
	}
	if(outl > 0) return pad0(out, outl); else return "";
}
/*jshint +W086 */
function commaify(s) {
	var w = 3;
	if(s.length <= w) return s;
	var j = (s.length % w), o = s.substr(0,j);
	for(; j!=s.length; j+=w) o+=(o.length > 0 ? "," : "") + s.substr(j,w);
	return o;
}
var write_num = (function make_write_num(){
var pct1 = /%/g;
function write_num_pct(type, fmt, val){
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
function write_num_cm(type, fmt, val){
	var idx = fmt.length - 1;
	while(fmt.charCodeAt(idx-1) === 44) --idx;
	return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
function write_num_exp(fmt, val){
	var o;
	var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if(fmt.match(/^#+0.0E\+0$/)) {
		if(val == 0) return "0.0E+0";
		else if(val < 0) return "-" + write_num_exp(fmt, -val);
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
		var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
		if(ee < 0) ee += period;
		o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
		if(o.indexOf("e") === -1) {
			var fakee = Math.floor(Math.log(val)*Math.LOG10E);
			if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
			else o += "E+" + (fakee - ee);
			while(o.substr(0,2) === "0.") {
				o = o.charAt(0) + o.substr(2,period) + "." + o.substr(2+period);
				o = o.replace(/^0+([1-9])/,"$1").replace(/^0+\./,"0.");
			}
			o = o.replace(/\+-/,"-");
		}
		o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
	} else o = val.toExponential(idx);
	if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
	if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
	return o.replace("e","E");
}
var frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
function write_num_f1(r, aval, sign) {
	var den = parseInt(r[4],10), rr = Math.round(aval * den), base = Math.floor(rr/den);
	var myn = (rr - base*den), myd = den;
	return sign + (base === 0 ? "" : ""+base) + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad_(myn,r[1].length) + r[2] + "/" + r[3] + pad0(myd,r[4].length));
}
function write_num_f2(r, aval, sign) {
	return sign + (aval === 0 ? "" : ""+aval) + fill(" ", r[1].length + 2 + r[4].length);
}
var dec1 = /^#*0*\.([0#]+)/;
var closeparen = /\).*[0#]/;
var phone = /\(###\) ###\\?-####/;
function hashq(str) {
	var o = "", cc;
	for(var i = 0; i != str.length; ++i) switch((cc=str.charCodeAt(i))) {
		case 35: break;
		case 63: o+= " "; break;
		case 48: o+= "0"; break;
		default: o+= String.fromCharCode(cc);
	}
	return o;
}
function rnd(val, d) { var dd = Math.pow(10,d); return ""+(Math.round(val * dd)/dd); }
function dec(val, d) {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 0;
	}
	return Math.round((val-Math.floor(val))*Math.pow(10,d));
}
function carry(val, d) {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 1;
	}
	return 0;
}
function flr(val) { if(val < 2147483647 && val > -2147483648) return ""+(val >= 0 ? (val|0) : (val-1|0)); return ""+Math.floor(val); }
function write_num_flt(type, fmt, val) {
	if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num_flt('n', ffmt, val);
		return '(' + write_num_flt('n', ffmt, -val) + ')';
	}
	if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
	if(fmt.indexOf('%') !== -1) return write_num_pct(type, fmt, val);
	if(fmt.indexOf('E') !== -1) return write_num_exp(fmt, val);
	if(fmt.charCodeAt(0) === 36) return "$"+write_num_flt(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
	var o;
	var r, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
	if(fmt.match(/^00+$/)) return sign + pad0r(aval,fmt.length);
	if(fmt.match(/^[#?]+$/)) {
		o = pad0r(val,0); if(o === "0") o = "";
		return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(frac1))) return write_num_f1(r, aval, sign);
	if(fmt.match(/^#+0+$/)) return sign + pad0r(aval,fmt.length - fmt.indexOf("0"));
	if((r = fmt.match(dec1))) {
		o = rnd(val, r[1].length).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1])).replace(/\.(\d*)$/,function($$, $1) { return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		return sign + rnd(aval, r[2].length).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify(pad0r(aval,0));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(""+(Math.floor(val) + carry(val, r[1].length))) + "." + pad0(dec(val, r[1].length),r[1].length);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num_flt(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g,""), val));
		ri = 0;
		return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
	}
	if(fmt.match(phone)) {
		o = write_num_flt(type, "##########", val);
		return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
	}
	var oa = "";
	if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length,7);
		ff = frac(aval, Math.pow(10,ri)-1, false);
		o = "" + sign;
		oa = write_num("n", r[1], ff[1]);
		if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
		o += oa + r[2] + "/" + r[3];
		oa = rpad_(ff[2],ri);
		if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
		o += oa;
		return o;
	}
	if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,ri)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^[#0?]+$/))) {
		o = pad0r(val, 0);
		if(fmt.length <= o.length) return o;
		return hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
		o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
		ri = o.indexOf(".");
		var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
		return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		ri = dec(val, r[1].length);
		return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(flr(val)).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(ri,r[1].length);
	}
	switch(fmt) {
		case "###,##0.00": return write_num_flt(type, "#,##0.00", val);
		case "###,###":
		case "##,###":
		case "#,###": var x = commaify(pad0r(aval,0)); return x !== "0" ? sign + x : "";
		case "###,###.00": return write_num_flt(type, "###,##0.00",val).replace(/^0\./,".");
		case "#,###.00": return write_num_flt(type, "#,##0.00",val).replace(/^0\./,".");
		default:
	}
	throw new Error("unsupported format |" + fmt + "|");
}
function write_num_cm2(type, fmt, val){
	var idx = fmt.length - 1;
	while(fmt.charCodeAt(idx-1) === 44) --idx;
	return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
function write_num_pct2(type, fmt, val){
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
function write_num_exp2(fmt, val){
	var o;
	var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if(fmt.match(/^#+0.0E\+0$/)) {
		if(val == 0) return "0.0E+0";
		else if(val < 0) return "-" + write_num_exp2(fmt, -val);
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
		var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
		if(ee < 0) ee += period;
		o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
		if(!o.match(/[Ee]/)) {
			var fakee = Math.floor(Math.log(val)*Math.LOG10E);
			if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
			else o += "E+" + (fakee - ee);
			o = o.replace(/\+-/,"-");
		}
		o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
	} else o = val.toExponential(idx);
	if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
	if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
	return o.replace("e","E");
}
function write_num_int(type, fmt, val) {
	if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num_int('n', ffmt, val);
		return '(' + write_num_int('n', ffmt, -val) + ')';
	}
	if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm2(type, fmt, val);
	if(fmt.indexOf('%') !== -1) return write_num_pct2(type, fmt, val);
	if(fmt.indexOf('E') !== -1) return write_num_exp2(fmt, val);
	if(fmt.charCodeAt(0) === 36) return "$"+write_num_int(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
	var o;
	var r, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
	if(fmt.match(/^00+$/)) return sign + pad0(aval,fmt.length);
	if(fmt.match(/^[#?]+$/)) {
		o = (""+val); if(val === 0) o = "";
		return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
	if(fmt.match(/^#+0+$/)) return sign + pad0(aval,fmt.length - fmt.indexOf("0"));
	if((r = fmt.match(dec1))) {
o = (""+val).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1]));
		o = o.replace(/\.(\d*)$/,function($$, $1) {
return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		return sign + (""+aval).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify((""+aval));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify((""+val)) + "." + fill('0',r[1].length);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num_int(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g,""), val));
		ri = 0;
		return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
	}
	if(fmt.match(phone)) {
		o = write_num_int(type, "##########", val);
		return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
	}
	var oa = "";
	if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length,7);
		ff = frac(aval, Math.pow(10,ri)-1, false);
		o = "" + sign;
		oa = write_num("n", r[1], ff[1]);
		if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
		o += oa + r[2] + "/" + r[3];
		oa = rpad_(ff[2],ri);
		if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
		o += oa;
		return o;
	}
	if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,ri)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^[#0?]+$/))) {
		o = "" + val;
		if(fmt.length <= o.length) return o;
		return hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
		o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
		ri = o.indexOf(".");
		var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
		return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify(""+val).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(0,r[1].length);
	}
	switch(fmt) {
		case "###,###":
		case "##,###":
		case "#,###": var x = commaify(""+aval); return x !== "0" ? sign + x : "";
		default:
			if(fmt.match(/\.[0#?]*$/)) return write_num_int(type, fmt.slice(0,fmt.lastIndexOf(".")), val) + hashq(fmt.slice(fmt.lastIndexOf(".")));
	}
	throw new Error("unsupported format |" + fmt + "|");
}
return function write_num(type, fmt, val) {
	return (val|0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
};})();
function split_fmt(fmt) {
	var out = [];
	var in_str = false/*, cc*/;
	for(var i = 0, j = 0; i < fmt.length; ++i) switch((/*cc=*/fmt.charCodeAt(i))) {
		case 34: /* '"' */
			in_str = !in_str; break;
		case 95: case 42: case 92: /* '_' '*' '\\' */
			++i; break;
		case 59: /* ';' */
			out[out.length] = fmt.substr(j,i-j);
			j = i+1;
	}
	out[out.length] = fmt.substr(j);
	if(in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
	return out;
}
SSF._split = split_fmt;
var abstime = /\[[HhMmSs]*\]/;
function fmt_is_date(fmt) {
	var i = 0, /*cc = 0,*/ c = "", o = "";
	while(i < fmt.length) {
		switch((c = fmt.charAt(i))) {
			case 'G': if(isgeneral(fmt, i)) i+= 6; i++; break;
			case '"': for(;(/*cc=*/fmt.charCodeAt(++i)) !== 34 && i < fmt.length;){/*empty*/} ++i; break;
			case '\\': i+=2; break;
			case '_': i+=2; break;
			case '@': ++i; break;
			case 'B': case 'b':
				if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") return true;
				/* falls through */
			case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
				/* falls through */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g': return true;
			case 'A': case 'a':
				if(fmt.substr(i, 3).toUpperCase() === "A/P") return true;
				if(fmt.substr(i, 5).toUpperCase() === "AM/PM") return true;
				++i; break;
			case '[':
				o = c;
				while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
				if(o.match(abstime)) return true;
				break;
			case '.':
				/* falls through */
			case '0': case '#':
				while(i < fmt.length && ("0#?.,E+-%".indexOf(c=fmt.charAt(++i)) > -1 || (c=='\\' && fmt.charAt(i+1) == "-" && "0#".indexOf(fmt.charAt(i+2))>-1))){/* empty */}
				break;
			case '?': while(fmt.charAt(++i) === c){/* empty */} break;
			case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break;
			case '(': case ')': ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1){/* empty */} break;
			case ' ': ++i; break;
			default: ++i; break;
		}
	}
	return false;
}
SSF.is_date = fmt_is_date;
function eval_fmt(fmt, v, opts, flen) {
	var out = [], o = "", i = 0, c = "", lst='t', dt, j, cc;
	var hr='H';
	/* Tokenize */
	while(i < fmt.length) {
		switch((c = fmt.charAt(i))) {
			case 'G': /* General */
				if(!isgeneral(fmt, i)) throw new Error('unrecognized character ' + c + ' in ' +fmt);
				out[out.length] = {t:'G', v:'General'}; i+=7; break;
			case '"': /* Literal text */
				for(o="";(cc=fmt.charCodeAt(++i)) !== 34 && i < fmt.length;) o += String.fromCharCode(cc);
				out[out.length] = {t:'t', v:o}; ++i; break;
			case '\\': var w = fmt.charAt(++i), t = (w === "(" || w === ")") ? w : 't';
				out[out.length] = {t:t, v:w}; ++i; break;
			case '_': out[out.length] = {t:'t', v:" "}; i+=2; break;
			case '@': /* Text Placeholder */
				out[out.length] = {t:'T', v:v}; ++i; break;
			case 'B': case 'b':
				if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") {
					if(dt==null) { dt=parse_date_code(v, opts, fmt.charAt(i+1) === "2"); if(dt==null) return ""; }
					out[out.length] = {t:'X', v:fmt.substr(i,2)}; lst = c; i+=2; break;
				}
				/* falls through */
			case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
				c = c.toLowerCase();
				/* falls through */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g':
				if(v < 0) return "";
				if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
				o = c; while(++i < fmt.length && fmt.charAt(i).toLowerCase() === c) o+=c;
				if(c === 'm' && lst.toLowerCase() === 'h') c = 'M';
				if(c === 'h') c = hr;
				out[out.length] = {t:c, v:o}; lst = c; break;
			case 'A': case 'a':
				var q={t:c, v:c};
				if(dt==null) dt=parse_date_code(v, opts);
				if(fmt.substr(i, 3).toUpperCase() === "A/P") { if(dt!=null) q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
				else if(fmt.substr(i,5).toUpperCase() === "AM/PM") { if(dt!=null) q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
				else { q.t = "t"; ++i; }
				if(dt==null && q.t === 'T') return "";
				out[out.length] = q; lst = c; break;
			case '[':
				o = c;
				while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
				if(o.slice(-1) !== ']') throw 'unterminated "[" block: |' + o + '|';
				if(o.match(abstime)) {
					if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
					out[out.length] = {t:'Z', v:o.toLowerCase()};
					lst = o.charAt(1);
				} else if(o.indexOf("$") > -1) {
					o = (o.match(/\$([^-\[\]]*)/)||[])[1]||"$";
					if(!fmt_is_date(fmt)) out[out.length] = {t:'t',v:o};
				}
				break;
			/* Numbers */
			case '.':
				if(dt != null) {
					o = c; while(++i < fmt.length && (c=fmt.charAt(i)) === "0") o += c;
					out[out.length] = {t:'s', v:o}; break;
				}
				/* falls through */
			case '0': case '#':
				o = c; while(++i < fmt.length && "0#?.,E+-%".indexOf(c=fmt.charAt(i)) > -1) o += c;
				out[out.length] = {t:'n', v:o}; break;
			case '?':
				o = c; while(fmt.charAt(++i) === c) o+=c;
				out[out.length] = {t:c, v:o}; lst = c; break;
			case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break; // **
			case '(': case ')': out[out.length] = {t:(flen===1?'t':c), v:c}; ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				o = c; while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) o+=fmt.charAt(i);
				out[out.length] = {t:'D', v:o}; break;
			case ' ': out[out.length] = {t:c, v:c}; ++i; break;
			case "$": out[out.length] = {t:'t', v:'$'}; ++i; break;
			default:
				if(",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(c) === -1) throw new Error('unrecognized character ' + c + ' in ' + fmt);
				out[out.length] = {t:'t', v:c}; ++i; break;
		}
	}
	var bt = 0, ss0 = 0, ssm;
	for(i=out.length-1, lst='t'; i >= 0; --i) {
		switch(out[i].t) {
			case 'h': case 'H': out[i].t = hr; lst='h'; if(bt < 1) bt = 1; break;
			case 's':
				if((ssm=out[i].v.match(/\.0+$/))) ss0=Math.max(ss0,ssm[0].length-1);
				if(bt < 3) bt = 3;
			/* falls through */
			case 'd': case 'y': case 'M': case 'e': lst=out[i].t; break;
			case 'm': if(lst === 's') { out[i].t = 'M'; if(bt < 2) bt = 2; } break;
			case 'X': /*if(out[i].v === "B2");*/
				break;
			case 'Z':
				if(bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
				if(bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
				if(bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
		}
	}
	switch(bt) {
		case 0: break;
		case 1:
if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
			if(dt.S >=  60) { dt.S = 0; ++dt.M; }
			if(dt.M >=  60) { dt.M = 0; ++dt.H; }
			break;
		case 2:
if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
			if(dt.S >=  60) { dt.S = 0; ++dt.M; }
			break;
	}
	/* replace fields */
	var nstr = "", jj;
	for(i=0; i < out.length; ++i) {
		switch(out[i].t) {
			case 't': case 'T': case ' ': case 'D': break;
			case 'X': out[i].v = ""; out[i].t = ";"; break;
			case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'b': case 'Z':
out[i].v = write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
				out[i].t = 't'; break;
			case 'n': case '?':
				jj = i+1;
				while(out[jj] != null && (
					(c=out[jj].t) === "?" || c === "D" ||
					((c === " " || c === "t") && out[jj+1] != null && (out[jj+1].t === '?' || out[jj+1].t === "t" && out[jj+1].v === '/')) ||
					(out[i].t === '(' && (c === ' ' || c === 'n' || c === ')')) ||
					(c === 't' && (out[jj].v === '/' || out[jj].v === ' ' && out[jj+1] != null && out[jj+1].t == '?'))
				)) {
					out[i].v += out[jj].v;
					out[jj] = {v:"", t:";"}; ++jj;
				}
				nstr += out[i].v;
				i = jj-1; break;
			case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
		}
	}
	var vv = "", myv, ostr;
	if(nstr.length > 0) {
		if(nstr.charCodeAt(0) == 40) /* '(' */ {
			myv = (v<0&&nstr.charCodeAt(0) === 45 ? -v : v);
			ostr = write_num('n', nstr, myv);
		} else {
			myv = (v<0 && flen > 1 ? -v : v);
			ostr = write_num('n', nstr, myv);
			if(myv < 0 && out[0] && out[0].t == 't') {
				ostr = ostr.substr(1);
				out[0].v = "-" + out[0].v;
			}
		}
		jj=ostr.length-1;
		var decpt = out.length;
		for(i=0; i < out.length; ++i) if(out[i] != null && out[i].t != 't' && out[i].v.indexOf(".") > -1) { decpt = i; break; }
		var lasti=out.length;
		if(decpt === out.length && ostr.indexOf("E") === -1) {
			for(i=out.length-1; i>= 0;--i) {
				if(out[i] == null || 'n?'.indexOf(out[i].t) === -1) continue;
				if(jj>=out[i].v.length-1) { jj -= out[i].v.length; out[i].v = ostr.substr(jj+1, out[i].v.length); }
				else if(jj < 0) out[i].v = "";
				else { out[i].v = ostr.substr(0, jj+1); jj = -1; }
				out[i].t = 't';
				lasti = i;
			}
			if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
		}
		else if(decpt !== out.length && ostr.indexOf("E") === -1) {
			jj = ostr.indexOf(".")-1;
			for(i=decpt; i>= 0; --i) {
				if(out[i] == null || 'n?'.indexOf(out[i].t) === -1) continue;
				j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")-1:out[i].v.length-1;
				vv = out[i].v.substr(j+1);
				for(; j>=0; --j) {
					if(jj>=0 && (out[i].v.charAt(j) === "0" || out[i].v.charAt(j) === "#")) vv = ostr.charAt(jj--) + vv;
				}
				out[i].v = vv;
				out[i].t = 't';
				lasti = i;
			}
			if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
			jj = ostr.indexOf(".")+1;
			for(i=decpt; i<out.length; ++i) {
				if(out[i] == null || ('n?('.indexOf(out[i].t) === -1 && i !== decpt)) continue;
				j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")+1:0;
				vv = out[i].v.substr(0,j);
				for(; j<out[i].v.length; ++j) {
					if(jj<ostr.length) vv += ostr.charAt(jj++);
				}
				out[i].v = vv;
				out[i].t = 't';
				lasti = i;
			}
		}
	}
	for(i=0; i<out.length; ++i) if(out[i] != null && 'n?'.indexOf(out[i].t)>-1) {
		myv = (flen >1 && v < 0 && i>0 && out[i-1].v === "-" ? -v:v);
		out[i].v = write_num(out[i].t, out[i].v, myv);
		out[i].t = 't';
	}
	var retval = "";
	for(i=0; i !== out.length; ++i) if(out[i] != null) retval += out[i].v;
	return retval;
}
SSF._eval = eval_fmt;
var cfregex = /\[[=<>]/;
var cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
function chkcond(v, rr) {
	if(rr == null) return false;
	var thresh = parseFloat(rr[2]);
	switch(rr[1]) {
		case "=":  if(v == thresh) return true; break;
		case ">":  if(v >  thresh) return true; break;
		case "<":  if(v <  thresh) return true; break;
		case "<>": if(v != thresh) return true; break;
		case ">=": if(v >= thresh) return true; break;
		case "<=": if(v <= thresh) return true; break;
	}
	return false;
}
function choose_fmt(f, v) {
	var fmt = split_fmt(f);
	var l = fmt.length, lat = fmt[l-1].indexOf("@");
	if(l<4 && lat>-1) --l;
	if(fmt.length > 4) throw new Error("cannot find right format for |" + fmt.join("|") + "|");
	if(typeof v !== "number") return [4, fmt.length === 4 || lat>-1?fmt[fmt.length-1]:"@"];
	switch(fmt.length) {
		case 1: fmt = lat>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
		case 2: fmt = lat>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
		case 3: fmt = lat>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
		case 4: break;
	}
	var ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
	if(fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) return [l, ff];
	if(fmt[0].match(cfregex) != null || fmt[1].match(cfregex) != null) {
		var m1 = fmt[0].match(cfregex2);
		var m2 = fmt[1].match(cfregex2);
		return chkcond(v, m1) ? [l, fmt[0]] : chkcond(v, m2) ? [l, fmt[1]] : [l, fmt[m1 != null && m2 != null ? 2 : 1]];
	}
	return [l, ff];
}
function format(fmt,v,o) {
	if(o == null) o = {};
	var sfmt = "";
	switch(typeof fmt) {
		case "string":
			if(fmt == "m/d/yy" && o.dateNF) sfmt = o.dateNF;
			else sfmt = fmt;
			break;
		case "number":
			if(fmt == 14 && o.dateNF) sfmt = o.dateNF;
			else sfmt = (o.table != null ? (o.table) : table_fmt)[fmt];
			break;
	}
	if(isgeneral(sfmt,0)) return general_fmt(v, o);
	if(v instanceof Date) v = datenum_local(v, o.date1904);
	var f = choose_fmt(sfmt, v);
	if(isgeneral(f[1])) return general_fmt(v, o);
	if(v === true) v = "TRUE"; else if(v === false) v = "FALSE";
	else if(v === "" || v == null) return "";
	return eval_fmt(f[1], v, o, f[0]);
}
function load_entry(fmt, idx) {
	if(typeof idx != 'number') {
		idx = +idx || -1;
for(var i = 0; i < 0x0188; ++i) {
if(table_fmt[i] == undefined) { if(idx < 0) idx = i; continue; }
			if(table_fmt[i] == fmt) { idx = i; break; }
		}
if(idx < 0) idx = 0x187;
	}
table_fmt[idx] = fmt;
	return idx;
}
SSF.load = load_entry;
SSF._table = table_fmt;
SSF.get_table = function get_table() { return table_fmt; };
SSF.load_table = function load_table(tbl) {
	for(var i=0; i!=0x0188; ++i)
		if(tbl[i] !== undefined) load_entry(tbl[i], i);
};
SSF.init_table = init_table;
SSF.format = format;
};
make_ssf(SSF);
/*global module */
if(typeof module !== 'undefined' && typeof DO_NOT_EXPORT_SSF === 'undefined') module.exports = SSF;