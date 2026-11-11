const fs=require('fs'),vm=require('vm'),assert=require('assert');
const base=__dirname+'/../src/';
const context=vm.createContext({console,Map,Set});
vm.runInContext(fs.readFileSync(base+'Config.js','utf8')+'\n'+fs.readFileSync(base+'KHO_NVL/FilmStockService.js','utf8'),context);
const c=context.CONFIG.COLUMNS,n=context.CONFIG.SHEET_NAMES;
const row=data=>({data,index:1});
let tx=[row({N_NVL_ID:'R1',Loai_GD:'CẤP NVL',So_KG:'60',Nguoi_Nhan:''}),row({N_NVL_ID:'R1',Loai_GD:'CẤP NVL',So_KG:'10',Nguoi_Nhan:''}),row({N_NVL_ID:'R1',Loai_GD:'TRẢ NVL',So_KG:'5',Nguoi_Nhan:'NV1'}),row({N_NVL_ID:'R1',Loai_GD:'TRẢ NVL',So_KG:'8',Nguoi_Nhan:''})];
const tables={
[n.FILM_LINES]:[row({DN_ID:'D1',PO_NVL_ID:'PO1',VL_ID:'V1',Loai_Nhua:'PET',Don_Gia:'20000',Ghi_Chu:'Ghi chú dòng'})],
[n.FILM_PO]:[row({PO_NVL_ID:'PO1',Ma_NCC:'S1',Ghi_Chu:'Ghi chú PO'})],
[n.FILM_SUPPLIERS]:[row({Ma_NCC:'S1',Ten_NCC:'NCC A'})],
[n.FILM_RECEIPTS]:[row({N_NVL_ID:'R1',DN_ID:'D1',Ngay_Nhan:'2026-11-11',Ngay_San_Xuat:'2026-11-10',So_KG_Nhan:'100',So_Cuon_Nhan:'2'})],
[n.PRODUCTS]:[row({SP_ID:'P1',name_napaco:'Sản phẩm A',type:'2'})],
[n.ORDER_DETAILS]:[row({DN_ID:'D1',Ma_SP:'P1'})]
};
context.sheetText_=v=>String(v??'').trim();context.filmStockAccess_=()=>({master:{},operation:{},permissions:new Map()});
context.readSheetTable_=(_book,names)=>({rows:names[0]===n.FILM_MOVEMENTS?tx:tables[names[0]]||[]});
context.filmMaterials_=()=>({rows:[row({VL_ID:'V1','Độ Dày':'0.8','Khổ Nhựa':'600'})]});
let result=context.getFilmStock();assert.equal(result.rows[0].stock,35);assert.equal(result.pendingCount,1);assert.equal(result.pendingKg,8);assert.equal(result.rows[0].products[0].type,'Bế');
assert.equal(context.getFilmStock({thicknessMin:0.9}).rows.length,0);
assert.throws(()=>context.getFilmStock({widthMin:700,widthMax:600}),/min/);
tx[3].data.Nguoi_Nhan='NV2';assert.equal(context.getFilmStock().rows[0].stock,43);assert.equal(context.filmStockMovements_({}).pendingCount,0);
tx[0].data.So_KG='NaN';assert.throws(()=>context.getFilmStock(),/không hợp lệ/);tx=[];
tables[n.FILM_RECEIPTS]=Array.from({length:1005},(_,i)=>({index:i,data:{N_NVL_ID:'R'+i,DN_ID:'D1',Ngay_Nhan:i===0?'2026-11-11':'2026-11-10',So_KG_Nhan:'0',So_Cuon_Nhan:'1'}}));
result=context.getFilmStock({status:'empty'});assert.equal(result.rows.length,1000);assert.equal(result.total,1005);assert.equal(result.rows[0].id,'R0');assert.equal(result.rows[1].id,'R1004');
assert.equal(context.getFilmStock().rows.length,0);
console.log('PASS: stock, confirmed/pending returns, filters, validation, product type and 1,000-row limit.');
