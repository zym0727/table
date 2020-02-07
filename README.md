# table

# 表格导出Excel小demo

# 操作前需要：   xlsx-style的bug修复:node_module/xlsx-style/dist/cpexcel.js的807行的var cpt = require('./cpt' + 'able')改为var cpt = cptable;

# 如果要设置行高，需要修改node_modules的xlsx-style文件下xlsx.js文件(不是dist文件下的xlsx.js文件)的write_ws_xml_data函数
# 删除这个函数，加上以下代码
var DEF_PPI = 96, PPI = DEF_PPI;
function px2pt(px) { return px * 96 / PPI; }
function pt2px(pt) { return pt * PPI / 96; }
function write_ws_xml_data(ws, opts, idx, wb) {
	var o = [], r = [], range = safe_decode_range(ws['!ref']), cell="", ref, rr = "", cols = [], R=0, C=0, rows = ws['!rows'];
	var dense = Array.isArray(ws);
	var params = ({r:rr}), row, height = -1;
	for(C = range.s.c; C <= range.e.c; ++C) cols[C] = encode_col(C);
	for(R = range.s.r; R <= range.e.r; ++R) {
		r = [];
		rr = encode_row(R);
		for(C = range.s.c; C <= range.e.c; ++C) {
			ref = cols[C] + rr;
			var _cell = dense ? (ws[R]||[])[C]: ws[ref];
			if(_cell === undefined) continue;
			if((cell = write_ws_xml_cell(_cell, ref, ws, opts, idx, wb)) != null) r.push(cell);
		}
		if(r.length > 0 || (rows && rows[R])) {
			params = ({r:rr});
			if(rows && rows[R]) {
				row = rows[R];
				if(row.hidden) params.hidden = 1;
				height = -1;
				if (row.hpx) height = px2pt(row.hpx);
				else if (row.hpt) height = row.hpt;
				if (height > -1) { params.ht = height; params.customHeight = 1; }
				if (row.level) { params.outlineLevel = row.level; }
			}
			o[o.length] = (writextag('row', r.join(""), params));
		}
	}
	if(rows) for(; R < rows.length; ++R) {
		if(rows && rows[R]) {
			params = ({r:R+1});
			row = rows[R];
			if(row.hidden) params.hidden = 1;
			height = -1;
			if (row.hpx) height = px2pt(row.hpx);
			else if (row.hpt) height = row.hpt;
			if (height > -1) { params.ht = height; params.customHeight = 1; }
			if (row.level) { params.outlineLevel = row.level; }
			o[o.length] = (writextag('row', "", params));
		}
	}
	return o.join("");
}