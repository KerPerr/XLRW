#ifndef _XLRW_XLRW_h_
#define _XLRW_XLRW_h_

#include <Core/Core.h>
#define STRINGIFY(...) #__VA_ARGS__

using namespace Upp;

class Workbook;
class Sheet;
class Cell;

class Functions
{
	public:
	int ltoi(Upp::String c)
	{
		int ret = 0;
		
		char split[c.GetLength()+1];
		strcpy(split, c.ToStd().c_str());
		
		for(int i=0;i<c.GetLength();i++)
	        ret += (int)c[i] - 64;
		
		return c.GetLength() > 1 ? ret + c.GetLength() * 26 - 27 : ret ;
	}
	
	Upp::String itol(int val)
	{
		int dividend = val;
	    Upp::String col = "";
	    int modulo = 0;
	
	    while (dividend > 0) {
	        modulo = (dividend - 1) % 26;
	        col = Upp::AsString(char(65 + modulo)) + col;
	        dividend = (int)((dividend - modulo) / 26);
	    } 
	
	    return col;
	}
};

class Workbook : public Functions
{
	private:
		Upp::String file;
	public:
		Workbook(Upp::String filePath);
		~Workbook();
		
		Upp::VectorMap<Upp::String, Upp::String> files;
		Upp::Vector<Upp::String> values;
		Upp::Vector<Sheet> sheets;
		
		Sheet& sheet(int index);
		Sheet& sheet(Upp::String name);
		Sheet& AddSheet(Upp::String name);
		
		int GetIndex(Upp::String);
		void Update();
		
		void Save();
};

class Sheet : public Upp::Moveable<Sheet>, Functions
{
	private:
		int index;
		Workbook* parent;
		Upp::String name;
		Upp::String content;
	public:
		Sheet();
		Sheet(const Sheet& ws);
		Sheet(Workbook* wb, int index, Upp::String name, Upp::String content);
		Sheet& operator=(const Sheet& ws);
		~Sheet();
		
		Upp::String GetContent() const;
		Upp::String GetName() const;
		int GetIndex() const;
		
		int lastRow();
		int lastCol();
		
		Upp::Vector<Cell> cells;
		Cell& cell(int row, int col);
		Cell& cell(int row, Upp::String col);
		
		void Save();
};

class Cell : public Upp::Moveable<Cell>, Functions
{
	private:
		Sheet* parent;
		Upp::String value;
	public:
		int row = 0;
		int col = 0;
		
		Cell();
		Cell(int row, int col, Upp::String value);
		~Cell();

		Upp::String Value();
		
		void Value(Upp::String val);
		void Value(int val);
		
		void setParent(Sheet* ws);
};

#endif
