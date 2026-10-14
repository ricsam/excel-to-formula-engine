import { expect, test } from "bun:test";
import { zipSync, strToU8 } from "fflate";
import { FormulaEngine, parseCellReference } from "@ricsam/formula-engine";
import { excelToFormulaEngine } from "../../index";
import { readXlsxPackage } from "../../src/xlsx/package";
import { parseStyles } from "../../src/xlsx/styles";

// Hand-authored OOXML is intentional: higher-level writers often materialize
// inherited styles onto cells and would hide precisely these regressions.
const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
 <font><name val="Aptos"/><sz val="12"/><color rgb="FF000000"/></font>
 <font><name val="Times New Roman"/><b/><i/><u/><sz val="18"/><color rgb="FF008000"/></font>
 <font><name val="Arial"/><sz val="9"/><color rgb="FF0000FF"/></font>
</fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellStyleXfs>
<cellXfs count="6">
 <xf fontId="0" numFmtId="0" xfId="0"/>
 <xf fontId="2" xfId="0"/>
 <xf xfId="1"/>
 <xf fontId="2" xfId="1" applyFont="0"/>
 <xf fontId="2" xfId="1" applyFont="1"/>
 <xf fontId="0" numFmtId="14"/>
</cellXfs><dxfs count="1"><dxf><font><color rgb="FFFF0000"/></font></dxf></dxfs>
</styleSheet>`;

function fixture(sheet: string, styleXml: string | undefined = styles): Blob {
  const files: Record<string, string> = {
    "xl/workbook.xml": `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    "xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    "xl/worksheets/sheet1.xml": `<worksheet>${sheet}</worksheet>`,
  };
  if (styleXml) files["xl/styles.xml"] = styleXml;
  const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([k,v])=>[k,strToU8(v)])));
  return new Blob([bytes as unknown as ArrayBuffer]);
}
const cell = (ref: string, attrs = "", value = "same") => `<c r="${ref}" t="inlineStr" ${attrs}><is><t>${value}</t></is></c>`;
async function imported(sheet: string, styleXml = styles) {
  const data = await excelToFormulaEngine(fixture(sheet, styleXml));
  const engine = FormulaEngine.buildEmpty();
  engine.addWorkbook({workbookName:"Test",data});
  const style = (ref: string) => engine.getCellStyle({workbookName:"Test",sheetName:"Sheet1",...parseCellReference(ref)});
  return {data,engine,style};
}

test("missing cell s inherits XF0, matching explicitly styled neighbor", async () => {
  const {data,style}=await imported(`<sheetData><row r="13">${cell("B13",'s="0"')}</row><row r="14">${cell("B14")}</row></sheetData>`);
  expect(style("B14")).toEqual(style("B13"));
  expect(style("B14")).toMatchObject({fontSize:16,color:"#000000"});
  expect(data.workbookMetadata?.defaultCellStyle).toEqual(style("B14"));
});

test("cell > active row > column > default, including explicit s=0", async () => {
  const {style}=await imported(`<cols><col min="2" max="16384" style="1"/></cols><sheetData>
  <row r="1">${cell("A1")}${cell("B1")}</row>
  <row r="2" s="2" customFormat="1">${cell("B2")}${cell("C2",'s="0"')}${cell("D2",'s="1"')}</row>
  <row r="3" s="2" customFormat="0">${cell("B3")}</row>
  <row r="4" s="2">${cell("B4")}</row></sheetData>`);
  expect(style("A1")?.fontSize).toBe(16);
  expect(style("B1")?.fontSize).toBe(12);
  expect(style("B2")).toMatchObject({fontSize:24,bold:true,color:"#008000"});
  expect(style("C2")?.fontSize).toBe(16);
  expect(style("D2")?.fontSize).toBe(12);
  expect(style("B3")?.fontSize).toBe(12);
  expect(style("B4")?.fontSize).toBe(12);
});

test("named-style font inherits or replaces atomically, respecting applyFont", async () => {
  const {style}=await imported(`<sheetData><row r="1">${cell("A1",'s="2"')}${cell("B1",'s="3"')}${cell("C1",'s="4"')}${cell("D1",'s="1"')}</row></sheetData>`);
  expect(style("A1")).toEqual(style("B1"));
  expect(style("A1")).toMatchObject({bold:true,italic:true,underline:true,fontSize:24});
  expect(style("C1")).toEqual(style("D1"));
  expect(style("C1")).toMatchObject({bold:false,italic:false,underline:false,fontSize:12});
});

test("default XF can itself inherit a named font, and explicit plain font clears it", async () => {
  const namedDefault = styles.replace('<xf fontId="0" numFmtId="0" xfId="0"/>', '<xf xfId="1"/>');
  const {data,style}=await imported(`<sheetData><row r="1">${cell("A1")}${cell("B1",'s="4"')}</row></sheetData>`, namedDefault);
  expect(data.workbookMetadata?.defaultCellStyle).toMatchObject({bold:true,fontSize:24});
  expect({...data.workbookMetadata?.defaultCellStyle,...style("B1")}).toMatchObject({bold:false,italic:false,underline:false,fontSize:12});
});

test("styles:false opts out of both default metadata and direct styles", async () => {
  const data=await excelToFormulaEngine(fixture(`<sheetData><row r="1">${cell("A1")}</row></sheetData>`),{styles:false});
  expect(data.cellStyles).toBeUndefined();
  expect(data.workbookMetadata?.defaultCellStyle).toBeUndefined();
});

test("inherited number formats feed date conversion and metadata", async () => {
  const {data}=await imported(`<cols><col min="1" max="1" style="5"/></cols><sheetData><row r="1"><c r="A1"><v>45000</v></c></row></sheetData>`);
  expect((data.sheets[0]!.content as Map<string,unknown>).get("A1")).toBe("2023-03-15");
  expect((data.sheets[0]!.cellMetadata as Map<string,unknown>).get("A1")).toEqual({numberFormat:"mm-dd-yy"});
});

test("differential fonts stay partial rather than clearing decorations", async () => {
  const pkg = await readXlsxPackage(fixture("<sheetData/>"));
  expect(parseStyles(pkg).differentialFormats[0]).toEqual({color:"#ff0000"});
});

test("row and column inheritance retains fills, borders and alignment without merging XFs", async () => {
  const xml = styles
    .replace('<fills count="1"><fill><patternFill patternType="none"/></fill></fills>', '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>')
    .replace('<borders count="1"><border/></borders>', '<borders count="2"><border/><border><bottom style="thin"><color rgb="FF000000"/></bottom></border></borders>')
    .replace('<xf fontId="2" xfId="0"/>', '<xf fontId="2" xfId="0" fillId="1" borderId="1" applyFill="true" applyBorder="1" applyAlignment="1"><alignment wrapText="1"/></xf>');
  const { style } = await imported(`<cols><col min="1" max="3" style="1"/></cols><sheetData>
    <row r="1">${cell("A1")}${cell("B1", 's="0"')}</row>
    <row r="2" s="1" customFormat="1">${cell("D2")}</row>
    <row r="3" s="0" customFormat="1">${cell("A3")}</row></sheetData>`, xml);
  expect(style("A1")).toMatchObject({ backgroundColor: "#ffff00", wrapText: true, borderSides: { bottom: true } });
  expect(style("D2")).toEqual(style("A1"));
  for (const ref of ["B1", "A3"]) {
    expect(style(ref)?.backgroundColor).toBeUndefined();
    expect(style(ref)?.wrapText).toBeUndefined();
    expect(style(ref)?.borderSides).toBeUndefined();
  }
  const disabled = xml.replace('applyFill="true"', 'applyFill="false"')
    .replace('applyBorder="1"', 'applyBorder="0"').replace('applyAlignment="1"', 'applyAlignment="0"');
  const table = parseStyles(await readXlsxPackage(fixture("<sheetData/>", disabled)));
  expect(table.formats[1]?.style.backgroundColor).toBeUndefined();
  expect(table.formats[1]?.style.wrapText).toBeUndefined();
  expect(table.formats[1]?.style.borderSides).toBeUndefined();
});

test("a named style can exclude its font component", async () => {
  const xml = styles.replace('<xf fontId="1"/>', '<xf fontId="1" applyFont="false"/>');
  const table = parseStyles(await readXlsxPackage(fixture("<sheetData/>", xml)));
  expect(table.formats[2]?.style).toMatchObject({ fontFamily: "Aptos", bold: false });
  expect(table.formats[3]?.style).toMatchObject({ fontFamily: "Aptos", bold: false });
  expect(table.formats[4]?.style).toMatchObject({ fontFamily: "Arial", bold: false });
});

test("font names survive default, inherited and overridden fonts", async () => {
  const { data, style } = await imported(`<sheetData><row r="1">${cell("A1")}${cell("B1", 's="3"')}${cell("C1", 's="4"')}</row></sheetData>`);
  expect(data.workbookMetadata?.defaultCellStyle).toMatchObject({ fontFamily: "Aptos" });
  expect(style("B1")).toMatchObject({ fontFamily: "Times New Roman" });
  expect(style("C1")).toMatchObject({ fontFamily: "Arial" });
});

test("boolean spellings and row XF zero preserve precedence", async () => {
  const booleanStyles = styles.replace('applyFont="0"', 'applyFont="false"').replace('applyFont="1"', 'applyFont="true"');
  const { style } = await imported(`<cols><col min="1" max="3" style="1"/></cols><sheetData>
    <row r="1" s="0" customFormat="true">${cell("A1")}${cell("B1", 's="3"')}${cell("C1", 's="4"')}</row>
    <row r="2" s="2" customFormat="false">${cell("A2")}</row></sheetData>`, booleanStyles);
  expect(style("A1")).toMatchObject({ fontSize: 16 });
  expect(style("B1")).toMatchObject({ fontSize: 24 });
  expect(style("C1")).toMatchObject({ fontSize: 12 });
  expect(style("A2")).toMatchObject({ fontSize: 12 });
});

test("underline enums and explicit differential decoration resets are retained", async () => {
  const xml = styles.replace('<u/>', '<u val="doubleAccounting"/>')
    .replace('<dxf><font><color rgb="FFFF0000"/></font></dxf>', '<dxf><font><b val="false"/><i val="0"/><u val="none"/></font></dxf>');
  const table = parseStyles(await readXlsxPackage(fixture("<sheetData/>", xml)));
  expect(table.formats[2]?.style.underline).toBe(true);
  expect(table.differentialFormats[0]).toEqual({ bold: false, italic: false, underline: false });
});

test("full column defaults do not materialize absent coordinates", async () => {
  const { data } = await imported(`<cols><col min="1" max="16384" style="1"/></cols><sheetData><row r="1">${cell("A1")}</row></sheetData>`);
  expect((data.sheets[0]!.content as Map<string, unknown>).size).toBe(1);
});

test("missing styles part does not invent a workbook font", async () => {
  const data=await excelToFormulaEngine(fixture(`<sheetData><row r="1">${cell("A1")}</row></sheetData>`, ""));
  expect(data.cellStyles).toBeUndefined();
  expect(data.workbookMetadata?.defaultCellStyle).toBeUndefined();
});
