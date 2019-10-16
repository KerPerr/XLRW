#include <Core/Core.h>
#include <plugin/zip/zip.h>
#include <plugin/pcre/Pcre.h>
#include <chrono>
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
}

Sheet Workbook::sheet(Upp::String name)
{
	for(Sheet& sht : sheets) {
		if(sht.GetName() == name)
			return sht;
	}
}

Workbook::~Workbook(){}

Sheet::Sheet(){}

Sheet::Sheet(const Sheet& ws)
{
	index = ws.GetIndex();
	name = ws.GetName();
	content = ws.GetContent();
}

Sheet::Sheet(String name)
{
	this->name = name;
}

Sheet::Sheet(int index, String name)
{
	this->index = index;
	this->name = name;
}

Sheet::Sheet(int index, String name, String content)
{
	this->index = index;
	this->name = name;
	this->content = content;
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
/*
Cell Sheet::cell(int row, int col)
{
	for(Cell& c : cells) {
		
	}
}
*/
String	Sheet::GetContent()	{ return content; };
String	Sheet::GetName()	{ return name; };
int		Sheet::GetIndex()	{ return index; };

int	Sheet::lastRow()
{
	XmlNode xn = ParseXML(content);
	const XmlNode& nodes = xn["worksheet"]["dimension"];
	RegExp r1("([0-9]+)");
	String range = nodes.Attr("ref");
	
    while(r1.GlobalMatch(range)) {}
	if(r1.IsError())
	    Cout() << r1.GetError() << '\n';
	
	return stoi(r1[0].ToStd());
}

int	Sheet::lastCol()
{
	int ret = 0;
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
	
	char split[res.GetLength()+1];
	strcpy(split, res.ToStd().c_str());
	
	for(int i=0;i<res.GetLength();i++)
        ret += (int)res[i] - 64;
	
	return res.GetLength() > 1 ? ret + res.GetLength() * 26 - 27 : ret ;
}

Cell::Cell(){}
Cell::Cell(int row, int col, String value)
{
	this->row = row;
	this->col = col;
	this->value = value;
}
Cell::~Cell(){};

CONSOLE_APP_MAIN
{
	auto t1 = std::chrono::high_resolution_clock::now();
	
	Workbook wb("C:\\Users\\CASTREC\\Documents\\XML XL\\XML.xlsx");
	
	auto t2 = std::chrono::high_resolution_clock::now();
	auto duration = std::chrono::duration_cast<std::chrono::microseconds>( t2 - t1 ).count();
	Cout() << "Benchmark : " << duration << EOL;
	
	Sheet ws = wb.sheet(1);
	Cout() << "Name: " << ws.GetName() << EOL;
}