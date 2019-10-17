#include <Core/Core.h>
#include <plugin/zip/zip.h>
#include <plugin/pcre/Pcre.h>
#include <chrono>

#include <windows.h>
#include "XLRW.h"

using namespace Upp;

Workbook::Workbook(String filePath)
{
	FileUnZip unzip(filePath);
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

Workbook::~Workbook(){}

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
	Cout() << "Attr: " << nodes.Attr("ref") << EOL;
	RegExp r1("([0-9]+)");
	String range = nodes.Attr("ref");
	
    while(r1.GlobalMatch(range)) {}
	if(r1.IsError())
	    Cout() << r1.GetError() << '\n';
	
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
	    Cout() << r1.GetError() << '\n';
	
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
	
	Sheet ws = wb.sheet(4);
	Cout() << "Out: " << ws.cell(3, 3).Value() << EOL;
}