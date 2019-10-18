#include <Core/Core.h>
#include <plugin/zip/zip.h>
#include <plugin/pcre/Pcre.h>
#include <chrono>
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
	
	// Recupère les valeurs présentes dans les différentes sheets
	xn = ParseXML(files.Get("xl/sharedStrings.xml"));
	const XmlNode& val = xn["sst"];
	for(int i=0;i<val.GetCount();i++) {
		//Cout() << val[i]["t"][0].GetText() << EOL;
		values.Add(val[i]["t"][0].GetText());
	}
	
	// Recupère les noms des différentes sheets
	xn = ParseXML(files.Get("xl/workbook.xml"));
	const XmlNode& shts = xn["workbook"]["sheets"];
	for(int i=0;i<shts.GetCount();i++) {
		//Cout() << shts[i].Attr("name") << EOL;
		Sheet ws(i, shts[i].Attr("name"), files.Get("xl/worksheets/sheet"+AsString(i+1)+".xml"));
		sheets.Add(ws);
	}
}

Workbook::~Workbook()
{
	FileZip zip(file);
	
	for(int i=0;i<files.GetCount();i++){
		zip.WriteFile(files[i], files.GetKey(i));
	}
}

// Prévoire des exceptions si l'index ou le nom sont inexistant
Sheet Workbook::sheet(int index)
{
	for(Sheet& sht : sheets) {
		if(sht.GetIndex() == index)
			return sht;
	}
	MessageBox(0, "Feuille " + AsString(index) + " introuvable.", "Warning", MB_ICONWARNING | MB_OK);
	Sheet sht(-1);
	return sht;
}

Sheet Workbook::sheet(Upp::String name)
{
	for(Sheet& sht : sheets) {
		if(sht.GetName() == name)
			return sht;
	}
	MessageBox(0, "Feuille " + name + " introuvable.", "Warning", MB_ICONWARNING | MB_OK);
	Sheet sheet(-1);
	return sheet;
}

void Workbook::AddSheet(Upp::String name)
{
	// Je recupere l'ID le plus important
	int res = 0;
	XmlNode xn;
	RegExp rgx("([a-zA-Z]+)");
	xn = ParseXML(files.Get("xl/_rels/workbook.xml.rels"));
	const XmlNode& rss = xn["Relationships"];
	for(int i=0;i<rss.GetCount();i++) {
		String val = rss[i].Attr("Id");
		rgx.ReplaceGlobal(val, String(""));
		if(stoi(val.ToStd()) > res)
			res = stoi(val.ToStd());
	}
	
	// J'ajoute dans le fichiers des relations ma nouvelle feuille et je modifie les rId ...
	XmlNode& rel = xn("Relationships");
	
	int count = rel.GetCount();
	while(rel.GetCount()>0){
		rel.Remove(0);
	}
	
	for(int i=1;i<=sheets.GetCount()+1;i++){
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
	
	//J'ajoute dans le fichier workbook
	xn = ParseXML(files.Get("xl/workbook.xml"));
	XmlNode& ws = xn("workbook")("sheets").Add("sheet");
	ws.SetAttr("name", name);
	ws.SetAttr("sheetId", sheets.GetCount()+1);
	ws.SetAttr("r:id", "rId" + AsString(sheets.GetCount()+1));
	Cout() << AsXML(xn, XML_HEADER) << EOL;
	files.Get("xl/workbook.xml") = AsXML(xn, XML_HEADER);
	
	// Je crée le fichier xml.
	files.Add("xl/worksheets/sheet"+AsString(sheets.GetCount()+1)+".xml",
		#include "empty.xml"
	);
	
	// J'ajoute le fichier au vecteurs
	Sheet sht(sheets.GetCount()+1, name);
	sheets.Add(sht);
}

Sheet::Sheet(){}

Sheet::Sheet(int index){ this->index = index; }

Sheet::Sheet(const Sheet& ws)
{
	index = ws.GetIndex();
	name = ws.GetName();
	content = ws.GetContent();
	for(const Cell& c : ws.cells) { cells.Add(c); };
}

Sheet::Sheet(String name) { this->name = name; }

Sheet::Sheet(int index, String name)
{
	this->index = index;
	this->name = name;
}

Sheet::Sheet(int index, String name, String content)
{
	int r = 0;
	int c = 0;
	
	RegExp col("([0-9]+)");
	RegExp row("([A-Z]+)");
	String clear = "";
	
	this->index = index;
	this->name = name;
	this->content = content;
	
	XmlNode xn = ParseXML(files.Get("xl/worksheets/sheet"+AsString(index+1)+".xml"));
	const XmlNode& rows = xn["worksheet"]["sheetData"];
	for(int i=0;i<rows.GetCount();i++) {
		const XmlNode& nCells = rows[i];
		for(int j=0;j<nCells.GetCount();j++) {
			String cell = nCells[j].Attr("r");
			String outRow = cell;
			String outCol = cell;
			//Cout() << "Cell: " << cell << ", Value: " << values[stoi(nCells[j]["v"][0].GetText().ToStd())] << EOL;
			
			row.ReplaceGlobal(outRow, clear);
			r = stoi(outRow.ToStd());
			
			col.ReplaceGlobal(outCol, clear);
			c = ltoi(outCol);
			
			Cell out(r, c , values[stoi(nCells[j]["v"][0].GetText().ToStd())]);
			//Cout() << values[stoi(cells[j]["v"][0].GetText().ToStd())] << EOL;
			
			cells.Add(out);
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

Cell Sheet::cell(int row, int col)
{
	for(Cell& c : cells) {
		if(c.row == row && c.col == col)
			return c;
	}
	MessageBox(0, "Cellule introuvable.", "Warning", MB_ICONWARNING | MB_OK);
	Cell cell(0, 0, "");
	return cell;
}

String	Sheet::GetContent()	const	{ return content; };
String	Sheet::GetName()	const	{ return name; };
int		Sheet::GetIndex()	const	{ return index; };

int	Sheet::lastRow()
{
	XmlNode xn = ParseXML(content);
	const XmlNode& nodes = xn["worksheet"]["dimension"];
	//Cout() << "Attr: " << nodes.Attr("ref") << EOL;
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
	
    //while(r1.GlobalMatch(range)) {}
    while(r1.GlobalMatch(range)) {
		for(int i = 0; i < r1.GetCount(); i++)
			res = r1[i];
    }
	if(r1.IsError())
	    Cout() << r1.GetError() << EOL;
	
	return ltoi(res);
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

CONSOLE_APP_MAIN
{
	auto t1 = std::chrono::high_resolution_clock::now();
	
	Workbook wb("C:\\Users\\CASTREC\\Documents\\XML XL\\XML.xlsx");
	
	auto t2 = std::chrono::high_resolution_clock::now();
	auto duration = std::chrono::duration_cast<std::chrono::microseconds>( t2 - t1 ).count();
	Cout() << "Benchmark : " << duration << EOL;
	
	wb.AddSheet("LAST");
	
	/*
	Sheet ws = wb.sheet(4);
	Cout() << "Out: " << ws.cell(3, 3).Value() << EOL;
	*/
}