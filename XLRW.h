#ifndef _XLRW_XLRW_h_
#define _XLRW_XLRW_h_

#define STRINGIFY(...) #__VA_ARGS__

class Workbook;
class Sheet;
class Cell;

Upp::VectorMap<Upp::String, Upp::String> files;
Upp::Vector<Upp::String> values;

int ltoi(Upp::String c)
{
	int ret = 0;
	
	char split[c.GetLength()+1];
	strcpy(split, c.ToStd().c_str());
	
	for(int i=0;i<c.GetLength();i++)
        ret += (int)c[i] - 64;
	
	return c.GetLength() > 1 ? ret + c.GetLength() * 26 - 27 : ret ;
};

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

class Workbook
{
	private:
		Upp::String file;
	public:
		Workbook(Upp::String filePath);
		~Workbook();
		
		Upp::Vector<Sheet> sheets;
		
		Sheet sheet(int index);
		Sheet sheet(Upp::String name);
		
		void AddSheet(Upp::String name);
};

class Sheet : public Upp::Moveable<Sheet>
{
	private:
		int index;
		Upp::String name;
		Upp::String content;
	public:
		Sheet();
		Sheet(int index);
		Sheet(const Sheet& ws);
		Sheet(Upp::String name);
		Sheet(int index, Upp::String name);
		Sheet(int index, Upp::String name, Upp::String content);
		Sheet& operator=(const Sheet& ws);
		~Sheet();
		
		Upp::String GetContent() const;
		Upp::String GetName() const;
		int GetIndex() const;
		
		int lastRow();
		int lastCol();
		
		Upp::Vector<Cell> cells; // Error: Use of deleted function 'Sheet::Sheet(const Sheet&) -> Sans aucun probleme pour compiler
		Cell cell(int row, int col);
};

class Cell : public Upp::Moveable<Cell>
{
	private:
		Upp::String value;
	public:
		int row = 0;
		int col = 0;
		
		Cell();
		Cell(int row, int col, Upp::String value);
		~Cell();
		/*
		int GetRow();
		int GetCol();
		*/
		Upp::String Value();
		/*
		void Value(Upp::String val);
		void Value(int val);
		*/
};

#endif
