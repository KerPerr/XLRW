#ifndef _XLRW_XLRW_h_
#define _XLRW_XLRW_h_

class Workbook;
class Sheet;
class Cell;

class Workbook
{
	private:
		Upp::Vector<Upp::String> values;
		Upp::VectorMap<Upp::String, Upp::String> files;
	public:
		Workbook(Upp::String filePath);
		~Workbook();
		
		Upp::Vector<Sheet> sheets;
		
		Sheet sheet(int index);
		Sheet sheet(Upp::String name);
};

class Sheet : public Upp::Moveable<Sheet>
{
	private:
		int index;
		Upp::String name;
		Upp::String content;
	public:
		Sheet();
		Sheet(Upp::String name);
		Sheet(int index, Upp::String name);
		Sheet(int index, Upp::String name, Upp::String content);
		Sheet& operator=(Sheet& ws);
		~Sheet();
		
		Upp::Vector<Cell> cells; // Error: Use of deleted function 'Sheet::Sheet(const Sheet&) -> Sans aucun probleme pour compiler
		Upp::String GetContent();
		Upp::String GetName();
		int GetIndex();
		
		int lastRow();
		int lastCol();
};

class Cell : public Upp::Moveable<Cell>
{
	private:
		int row = 0;
		int col = 0;
		Upp::String value;
	public:
		Cell();
		Cell(int row, int col, Upp::String value);
		~Cell();
		/*
		int GetRow();
		int GetCol();
		
		Upp::String Value();
		void Value(Upp::String val);
		void Value(int val);
		*/
};

#endif
