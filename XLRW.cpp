#include <plugin/zip/zip.h>
#include <plugin/pcre/Pcre.h>
#include <fstream>

#include <windows.h>
#include "XLRW.h"

using namespace Upp;

Workbook::Workbook(String filePath)
{
	file = filePath;
	FileUnZip unzip(file);
	XmlNode xn;
		
	while(!(unzip.IsEof() || unzip.IsError())) {
		String file = unzip.GetPath();
		String content = unzip.ReadFile();
		files.Add(file, content);
	}
	
	xn = ParseXML(files.Get("xl/sharedStrings.xml"));
	const XmlNode& val = xn["sst"];
	for(int i=0;i<val.GetCount();i++) {
		values.Add(val[i]["t"][0].GetText());
	}
	
	xn = ParseXML(files.Get("xl/workbook.xml"));
	const XmlNode& xnws = xn["workbook"]["sheets"];
	for(int i=0;i<xnws.GetCount();i++) {
		sheets.Create(this, i, xnws[i].Attr("name"), files.Get("xl/worksheets/sheet"+AsString(i+1)+".xml"));
	}
}

Workbook::~Workbook()
{
	
}

Sheet& Workbook::sheet(int index)
{
	for(Sheet& ws : sheets) {
		if(ws.GetIndex() == index) {
			return ws;
		}
	}
	MessageBox(0, "Feuille " + AsString(index) + " introuvable.", "Warning", MB_ICONWARNING | MB_OK);
	throw std::exception();
}

Sheet& Workbook::sheet(Upp::String name)
{
	for(Sheet& ws : sheets) {
		if(ws.GetName().IsEqual(name)) {
			return ws;
		}
	}
	MessageBox(0, "Feuille " + name + " introuvable.", "Warning", MB_ICONWARNING | MB_OK);
	throw std::exception();
}

Sheet& Workbook::AddSheet(Upp::String name)
{
	for(Sheet& s : sheets) {
		if(s.GetName() == name) {
			MessageBox(0, "Feuille " + name + " existante.", "Warning", MB_ICONWARNING | MB_OK);
			return s;
		}
	}
	
	XmlNode xn = ParseXML(files.Get("xl/workbook.xml"));
	XmlNode& ws = xn("workbook")("sheets").Add("sheet");
	ws.SetAttr("name", name);
	ws.SetAttr("sheetId", sheets.GetCount()+1);
	ws.SetAttr("r:id", "rId" + AsString(sheets.GetCount()+1));
	//Cout() << AsXML(xn, XML_HEADER) << EOL;
	files.Get("xl/workbook.xml") = AsXML(xn, XML_HEADER);
	
	files.Add("xl/worksheets/sheet" + AsString(sheets.GetCount()+1) + ".xml",
		#include "empty.xml"
	);
	return sheets.Create(this, sheets.GetCount(), name,
		#include "empty.xml"
	);
}

int Workbook::GetIndex(String value)
{
	for(int i=0;i<values.GetCount();i++){
		if(values[i].IsEqual(value))
			return i;
	}
	
	if(value.GetCount() > 0) {
		values.Add(value);
		// Reconstuire sharedStrings.xml
		XmlNode xn = ParseXML(files.Get("xl/sharedStrings.xml"));
		XmlNode& data = xn("sst");

		while(data.GetCount() > 0){
			data.Remove(0);
		}
		
		for(String& s : values) {
			XmlNode& xnsi = data.Add("si");
			XmlNode& xnt = xnsi.Add("t");
			xnt.AddText(s);
		}
		files.Get("xl/sharedStrings.xml") = AsXML(xn, XML_HEADER);
		return values.GetCount()-1;
	}
	MessageBox(0, "GetIndex !", "Warning", MB_ICONWARNING | MB_OK);
	throw std::exception();
}

void Workbook::Update()
{
	XmlNode xn;
	xn = ParseXML(files.Get("xl/_rels/workbook.xml.rels"));
	XmlNode& rel = xn("Relationships");
	
	while(rel.GetCount() > 0){
		rel.Remove(0);
	}
	
	for(int i=1;i<=sheets.GetCount()+1;i++) {
		XmlNode& rl = rel.Add("Relationship");
		rl.SetAttr("Id", "rId" + AsString(i));
		rl.SetAttr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
		rl.SetAttr("Target", "worksheets/sheet"+ AsString(i) +".xml");
	}
	
	XmlNode& theme = rel.Add("Relationship");
	theme.SetAttr("Id", "rId" + AsString(sheets.GetCount()+2));
	theme.SetAttr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
	theme.SetAttr("Target", "theme/theme1.xml");
	
	XmlNode& style = rel.Add("Relationship");
	style.SetAttr("Id", "rId" + AsString(sheets.GetCount()+3));
	style.SetAttr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
	style.SetAttr("Target", "styles.xml");
	
	XmlNode& ss = rel.Add("Relationship");
	ss.SetAttr("Id", "rId" + AsString(sheets.GetCount()+4));
	ss.SetAttr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
	ss.SetAttr("Target", "sharedStrings.xml");
	
	//Cout() << AsXML(xn, XML_HEADER | XML_PRETTY) << EOL;
	
	files.Get("xl/_rels/workbook.xml.rels") = AsXML(xn, XML_HEADER);
}

void Workbook::Save()
{
	// Reconstruire la structure des fichiers a partir de chacun des vecteurs
	for(Sheet& ws : sheets) {
		ws.Save();
		files.Get("xl/worksheets/sheet" + AsString(ws.GetIndex()+1) + ".xml") = ws.GetContent();
	}
	Update();
	
	FileZip zip(file);
	
	for(int i=0;i<files.GetCount();i++){
		zip.WriteFile(files[i], files.GetKey(i));
	}
}

Sheet::Sheet(){}

Sheet::Sheet(const Sheet& ws)
{
	index = ws.GetIndex();
	name = ws.GetName();
	content = ws.GetContent();
	for(const Cell& c : ws.cells) { cells.Add(c); };
}

Sheet::Sheet(Workbook* wb, int index, String name, String content)
{
	int r = 0;
	int c = 0;
	parent = wb;
	
	RegExp col("([0-9]+)");
	RegExp row("([A-Z]+)");
	String clear = "";
	
	this->index = index;
	this->name = name;
	this->content = content;
	
	XmlNode xn = ParseXML(content);
	const XmlNode& xnr = xn["worksheet"]["sheetData"];
	for(int i=0;i<xnr.GetCount();i++) {
		const XmlNode& xnc = xnr[i];
		for(int j=0;j<xnc.GetCount();j++) {
			String cell = xnc[j].Attr("r");
			String outRow = cell;
			String outCol = cell;
			
			row.ReplaceGlobal(outRow, clear);
			r = stoi(outRow.ToStd());
			
			col.ReplaceGlobal(outCol, clear);
			c = ltoi(outCol);
			
			cells.Create(r, c , parent->values[stoi(xnc[j]["v"][0].GetText().ToStd())]);
		}
	}
}

Sheet& Sheet::operator=(const Sheet& ws)
{
	index = ws.GetIndex();
	name = ws.GetName();
	content = ws.GetContent();
	for(const Cell& c : ws.cells) {
		cells.Add(c);
	}
	return *this;
}

Sheet::~Sheet(){}

Cell& Sheet::cell(int row, int col)
{
	for(Cell& c : cells) {
		if(c.row == row && c.col == col) {
			return c;
		}
	}
	
	Cell& c = cells.Create(row, col, "");
	c.setParent(this);
	return c;
}

Cell& Sheet::cell(int row, String col)
{
	for(Cell& c : cells) {
		if(c.row == row && c.col == ltoi(col)) {
			return c;
		}
	}
	
	Cell& c = cells.Create(row, ltoi(col), "");
	c.setParent(this);
	return c;
}

String	Sheet::GetContent()	const	{ return content; };
String	Sheet::GetName()	const	{ return name; };
int		Sheet::GetIndex()	const	{ return index; };

int	Sheet::lastRow()
{
	XmlNode xn = ParseXML(content);
	const XmlNode& nodes = xn["worksheet"]["dimension"];
	
	RegExp r1("([0-9]+)");
	String range = nodes.Attr("ref");
	
    while(r1.GlobalMatch(range)) {}
	if(r1.IsError())
	    Cout() << r1.GetError() << EOL;
	
	return stoi(r1[0].ToStd());
}

int	Sheet::lastCol()
{
	String res = "";
	XmlNode xn = ParseXML(content);
	const XmlNode& nodes = xn["worksheet"]["dimension"];
	RegExp r1("([A-Z]+)");
	String range = nodes.Attr("ref");
	
    while(r1.GlobalMatch(range)) {
		for(int i = 0; i < r1.GetCount(); i++)
			res = r1[i];
    }
	if(r1.IsError())
	    Cout() << r1.GetError() << EOL;
	
	return ltoi(res);
}

void Sheet::Save()
{
	XmlNode xn = ParseXML(content);
	XmlNode& data = xn("worksheet")("sheetData");
	
	while(data.GetCount() > 0){
		data.Remove(0);
	}
		
	Sort(cells, [](const Cell& a, const Cell& b) { return ((a.row==b.row) ? a.col < b.col : a.row < b.row); });
	int current = 0;
	for(Cell& c : cells) {
		if(current != c.row) {
			current = c.row;
			XmlNode& xnr = data.Add("row");
			xnr.SetAttr("r", c.row);
			xnr.SetAttr("spans", "1:4");
			xnr.SetAttr("x14ac:dyDescent", "0.25");
			
			XmlNode& xnc = xnr.Add("c");
			xnc.SetAttr("r", itol(c.col) + AsString(c.row));
			xnc.SetAttr("t", "s");
			
			XmlNode& xnv = xnc.Add("v");
			
			xnv.AddText(AsString(parent->GetIndex(c.Value())));
		} else {
			int i = 0;
			for(const XmlNode& xnr : data) {
				if(xnr.Attr("r").IsEqual(AsString(c.row))) {
					XmlNode& xnc = data.At(i).Add("c");
					xnc.SetAttr("r", itol(c.col) + AsString(c.row));
					xnc.SetAttr("t", "s");
					XmlNode& xnv = xnc.Add("v");
					xnv.AddText(AsString(parent->GetIndex(c.Value())));
				}
				i++;
			}
		}
	}
	
	content = AsXML(xn, XML_HEADER);
	parent->files.Get("xl/worksheets/sheet" + AsString(GetIndex()+1) + ".xml") = content;
}

Cell::Cell(){}
Cell::Cell(int row, int col, String value)
{
	this->row = row;
	this->col = col;
	this->value = value;
}
Cell::~Cell(){};

String Cell::Value() { return value; };

void Cell::setParent(Sheet* ws) { this->parent = ws; }
void Cell::Value(String value) { this->value = value; }
void Cell::Value(int value) { this->value = AsString(value); }

/*
CONSOLE_APP_MAIN
{
	Workbook wb("C:\\Users\\CASTREC\\Documents\\XML XL\\XML.xlsx");
	
	wb.sheet(1).cell(1, "B").Value("CHANGE");
	wb.AddSheet("Quatrieme").cell(1, 1).Value("DID IT");
	wb.Save();
}
*/