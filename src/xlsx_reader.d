module xlsx_reader;

import std.algorithm;
import std.conv : to, parse;
import std.range : lockstep;
import std.stdio;
import std.string;
import std.metastrings;
import std.zip;
import std.stream;
import std.xml;
import std.variant;
import std.datetime;
import std.path;

template unit(string unitName) {
	const char[] helper = q{
		debug {
			write("%s.unittest ... ");
			scope(success) printf("passed\n");
			scope(failure) printf("NOT passed\n");
		}
	};
	
	const char[] unit = Format!(helper, unitName);
}

DateTime OleTimeToDateTime(double vtime) {
	return DateTime(1899, 12, 30) + seconds(cast(long)(vtime * 86400));
}

class ArchiveHandler {
	BufferedFile m_zip;
	string[string] m_xmlFiles;
	
	this(string filename) {
		m_zip = new BufferedFile(filename);
		char[] zipData = m_zip.readString(cast(uint)m_zip.size);
		ZipArchive archive = new ZipArchive(zipData);
		foreach (ArchiveMember am; archive.directory) {
			//writefln("member name is '%s'", am.name);
			m_xmlFiles[am.name] = cast(string)archive.expand(am);
		}
	}
	
	string[] getFileNames() {
		string[] xmlFileNames;
		foreach(name, data; m_xmlFiles) {
			xmlFileNames ~= name;
		}
		return xmlFileNames;
	}
	
	string getFile(string name) {
		return m_xmlFiles[name];
	}
	
	unittest {
		mixin( unit!("xlsx_reader.XlsxHandler") );
		ArchiveHandler handler = new ArchiveHandler("../../resource/excel_workbook.xlsx");
		string[] xmlFileNames = handler.getFileNames();
		assert( xmlFileNames.length != 0 );
	}
}

class Workbook {
	
	public {
		
		this(string filename) {
			m_archiveHandler = new ArchiveHandler(filename);
			
			// read shared string
			string sharedStringsXml = m_archiveHandler.getFile("xl/sharedStrings.xml");
			m_sharedStrings = readSharedString(sharedStringsXml);
			
			// read sheets
			string rawXmlWorkbook = m_archiveHandler.getFile("xl/workbook.xml");
			auto xmlWorkbook = new DocumentParser(rawXmlWorkbook);
			xmlWorkbook.onStartTag["sheet"] = (ElementParser xml) {
				uint sheetId = to!uint(xml.tag.attr["sheetId"]);
				string rawXmlSheet = m_archiveHandler.getFile("xl/worksheets/sheet" ~ to!string(sheetId) ~ ".xml");
				string sheetName = xml.tag.attr["name"];
				Sheet sheet = new Sheet(sheetId, sheetName, rawXmlSheet, m_sharedStrings);
				m_sheets ~= sheet;
			};
			xmlWorkbook.parse();
		}
		
		Sheet sheetByName(string sheetName) {
			foreach(sheet; m_sheets) {
				if ( sheet.name == sheetName ) return sheet;
			}
			assert(false, "sheet with name " ~ sheetName ~ " not exist");
		}
		
		immutable(string[]) getSharedStrings() pure {
			return cast(immutable)m_sharedStrings;
		}
		
	} // public
	
	private {
		
		string[] readSharedString(string sharedStringsXml) {
			string[] sharedStrings;
			auto sst = new DocumentParser(sharedStringsXml);
			
			uint count = to!uint(sst.tag.attr["count"]);
			sharedStrings = new string[](count);
			uint idx = 0;
			
			sst.onStartTag["si"] = (ElementParser si) {
				si.onEndTag["t"] = (in Element e) {
					sharedStrings[idx] = e.text();
					++idx;
				};
				si.parse();
			};
			sst.parse();
			
			return sharedStrings;
		}
		
	} // private
	
	private {
		
		ArchiveHandler m_archiveHandler;
		Sheet[] m_sheets;
		string[] m_sharedStrings;
		
	} // private
	
	unittest {
		mixin( unit!("xlsx_reader.Workbook") );
		Workbook workbook = new Workbook("../../resource/excel_workbook.xlsx");
	}
}

private {
	enum A2N = [ 'A':1, 'B':2, 'C':3, 'D':4, 'E':5, 'F':6,
				'G':7, 'H':8, 'I':9, 'J':10, 'K':11, 'L':12, 'M':13, 'N':14,
				'O':15, 'P':16, 'Q':17, 'R':18, 'S':19, 'T':20,
				'U':21, 'V':22, 'W':23, 'X':24, 'Y':25, 'Z':26 ];
	
	string nameByCol(ulong col, string _A2Z="ABCDEFGHIJKLMNOPQRSTUVWXYZ") {
		assert( col >= 0 );
		string name;
		while (true) {
			ulong quot = col / _A2Z.length;
			uint rem = col % cast(uint)_A2Z.length;
			name = _A2Z[rem] ~ name;
			if ( !quot ) return name;
			col = quot - 1;
		}
	}
	
	unittest {
		assert( nameByCol(0) == "A" );
		assert( nameByCol(1) == "B" );
		assert( nameByCol(26) == "AA" );
		assert( nameByCol(27) == "AB" );
	}
	
	uint colByName(string name) {
		assert( name.length > 0 );
		uint idx = reduce!("a * 26 + b - 'A' + 1")(0, name) - 1;
		return idx;
	}
	
	unittest {
		assert( colByName("A") == 0 );
		assert( colByName("B") == 1 );
		assert( colByName("AA") == 26 );
		assert( colByName("AB") == 27 );
	}
	
	auto cellByName(string name) {
		class Cell {
			uint row, col;
		}
		auto cell = new Cell;
		
		string strCol = munch(name, "A-Z");
		cell.row = to!uint(name) - 1;
		cell.col = colByName(strCol);
		
		return cell;
	}
}

class Sheet {
	uint m_id;
	string m_name;
	Cell[][] m_cells;
	string[] m_sharedStrings;
	class Merged { uint rlo, rhi, clo, chi; }
	Merged[] m_mergedCells;

	this(uint id, string name, string rawXmlSheet, ref string[] sharedStrings) {
		m_id = id;
		m_name = name;
		m_sharedStrings = sharedStrings;
		
		check(rawXmlSheet);
		auto xmlSheet = new DocumentParser(rawXmlSheet);
		xmlSheet.onStartTag["dimension"] = (ElementParser xmlRow) {
			string dimension = xmlRow.tag.attr["ref"];
			auto range = dimension.split(":");
			if ( range.length >= 2 ) {
				auto cellBegin = cellByName(range[0]);
				auto cellEnd = cellByName(range[1]);
				uint rows = cellEnd.row - cellBegin.row + 1;
				uint cols = cellEnd.col - cellBegin.col + 1;
				m_cells = new Cell[][](rows, cols);
			}
		};

		// merged cells
		xmlSheet.onStartTag["mergeCells"] = (ElementParser xmlMergedCells) {
			auto count = to!uint(xmlMergedCells.tag.attr["count"]);
			xmlMergedCells.onEndTag["mergeCell"] = (in Element mergedCell) {
				auto merged = mergedCell.tag.attr["ref"];
				auto range = merged.split(":");
				auto lo = cellByName(range[0]);
				auto hi = cellByName(range[1]);
				auto mergedInfo = new Merged;
				with(mergedInfo) {
					rlo = lo.row;
					clo = lo.col;
					rhi = hi.row;
					chi = hi.col;
				}
				m_mergedCells ~= mergedInfo;
			};
			xmlMergedCells.parse();
		};

		// values
		xmlSheet.onStartTag["row"] = (ElementParser xmlRow) {
			xmlRow.onStartTag["c"] = (ElementParser xmlCol) {
				auto cellname = xmlCol.tag.attr["r"];
				auto pos = cellByName(cellname);
				auto celltype = "t" in xmlCol.tag.attr;
				auto cellstyle = "s" in xmlCol.tag.attr;
				xmlCol.onEndTag["v"] = (in Element e) {
					if ( celltype !is null && *celltype == "s" ) { // shared string
						uint idx = to!uint(e.text());
						string value = to!string(m_sharedStrings[idx]);
						m_cells[pos.row][pos.col] = new Cell(value);
					}
					else if ( cellstyle !is null && *cellstyle == "1" ) { // currency
						double value = to!double(e.text());
						m_cells[pos.row][pos.col] = new Cell(value);
					}
					else if ( cellstyle !is null && *cellstyle == "2" ) { // datetime
						auto oledatetime = to!double(e.text());
						auto value = OleTimeToDateTime(oledatetime);
						m_cells[pos.row][pos.col] = new Cell(value);
					}
					else if ( celltype is null && cellstyle is null ) { // double
						auto value = to!double(e.text());
						m_cells[pos.row][pos.col] = new Cell(value);
					}
					else {
						auto value = e.text();
						writeln(cellname, " ", value, " unknown type");
						writefln("\tcelltype - %s, cellstyle - %s", celltype, cellstyle);
					}
				};
				xmlCol.parse();
			};
			xmlRow.parse();
		};
		xmlSheet.parse();
	}
	
	@property ulong id() {
		return m_id;
	}
	
	@property string name() {
		return m_name;
	}
	
	@property ulong rows() {
		//if( m_cells.empty ) return 0;
		return m_cells.length;
	}
	
	@property ulong cols() {
		//if( m_cells.empty || m_cells[0].empty ) return 0;
		return m_cells[0].length;
	}
	
	Cell[] row(uint idxRow)
	{
		return m_cells[idxRow];
	}

	Cell cell(string cellname) {
		auto pos = cellByName(cellname);
		return m_cells[pos.row][pos.col];
	}
	
	private {
		void printThis() {
			writeln("length: ", m_cells.sizeof);
			foreach(idxr, ro; m_cells) {
				foreach(idxc, co; ro) {
					writef(" %5s ", co.type);
				}
				write("\n");
			}		
		}

		bool isMerged(uint row, uint col) {
			foreach(range; m_mergedCells) {
				if(range.rlo <= row && range.rhi >= row && range.clo <= col && range.chi >= col) return true;
			}
			return false;
		}

		Cell mergedRef(uint row, uint col) {
			foreach(range; m_mergedCells) {
				if(range.rlo <= row && range.rhi >= row && range.clo <= col && range.chi >= col) return m_cells[range.rlo][range.clo];
			}
			return null;
		}
	}
	
	unittest {
		mixin( unit!("xlsx_reader.Sheet") );
		
		string[] strings;
		Sheet sheet = new Sheet(1, "Лист1", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook></workbook>", strings);
		assert( sheet.id == 1 );
		assert( sheet.name == "Лист1" );
	}
}

private class Range {
	Cell[][] m_cells;
	
public:
	
	this(size_t rows, size_t cols) {
		m_cells = new Cell[][](rows, cols);
	}
	
	void reserve(size_t rows, size_t cols) {
		m_cells = new Cell[][](rows, cols);
	}
	
	void opIndexAssign(Cell cell, size_t row, size_t col) {
		m_cells[row][col] = cell;
	}
	
	Cell opIndex(size_t row, size_t col) {
		return m_cells[row][col];
	}
	
	Range range(size_t rowBegin, size_t colBegin, size_t rowEnd, size_t colEnd)
	{
		size_t rows = rowEnd - rowBegin;
		size_t cols = colEnd - colBegin;
		Range range = new Range(0, 0);
		//foreach(idx, row; m_cells[2..4]) {
		//	range.m_cells ~= row[1..3];
		foreach(idx, row; m_cells[colBegin..colEnd]) {
			range.m_cells ~= row[rowBegin..rowEnd];
		}
		return range;
	}
	
	//void fun(T)(T value) if(is(typeof(T.type) : Cell)) {
	//	writeln("type Cell");
	//	writeln(value);
	//}
	
	Range range(string str)
	{
		writeln("old range");
		printThis();
		assert( m_cells != null );
		
		uint col, row;
		Range range;
		//writeln(str);
		
		auto r = str.split(":");
		if ( r.length == 1 ) {
			string strcol = munch(r[0], "A-Z");
			col = to!uint(colByName(strcol)) + 1;
			row = to!uint(r[0]);
			
			writeln("sliced range");
			//range = new Range(row, col);
			//range = this.range();
			//range.printThis();
			Variant range_variant;
			Cell cs = new Cell(3);
			range_variant = cs;
			//writeln(typeid(range_variant.type));
			//fun(range_variant);
		}
		
		//uint rr = 1;
		//uint cc = 0;
		//int[][] array = new int[][](rr, cc);
		//writeln("length ", array.length);
		//writeln("sizeof ", array.sizeof);
		//writeln("sizeof ", (uint[]).sizeof);
		//foreach(idxr, ro; array) {
		//	foreach(idxc, co; ro) {
		//		writeln(idxr, " ", idxc, " ",  co);
		//		//break;
		//	}
		//}
		////writeln(array);
		
		//else if ( r.length == 2 ) {
		//	string strRowMin = r[0];
		//	string strRowMax = r[1];
		//	string strColMin = munch(strRowMin, "A-Z");
		//	string strColMax = munch(strRowMax, "A-Z");
		//	uint rowMin = to!uint(strRowMin);
		//	uint rowMax = to!uint(strRowMax);
		//	uint colMin = colByName(strColMin);
		//	uint colMax = colByName(strColMax);
		//	uint rows = rowMax - rowMin;
		//	uint cols = colMax - colMin;
		//	m_cells = new string[][](rows, cols);
		//}
		
		//Range range = new Range(row, col);
		return range;
	}
	
	private void printThis()
	{
		writeln("length: ", m_cells.sizeof);
		foreach(idxr, ro; m_cells) {
			foreach(idxc, co; ro) {
				writef(" %5s ", co.value);
				//writef("%5s", idxc*idxr*1);
			}
			write("\n");
		}		
	}
	
	unittest {
		//Range range = new Range(8, 4);
		////range.reserve(3, 3);
		//foreach(row; [0, 1, 2, 3, 4, 5, 6, 7]){
		//	foreach(col; [0, 1, 2, 3]){
		//		Cell cell = new Cell(row + col);
		//		range[row, col] = cell;
		//		assert( range[row, col] == cell );
		//	}
		//}
		
		//Range range1 = range.range(1, 1, 3, 3);
		//range.range("A1");
	}
}

class Cell {
	Variant m_value;
	bool m_merged;
	Cell m_refToValue;
	
	this(T)(T value) {
		m_value = value;
	}

	this(T)(T value, ref Cell refToValue) {
		this(value);
		m_refToValue = refToValue;
		m_merged = true;
	}
	
	this(T : Variant)(T value) {
		m_value = value.get!value.type;
	}
	
	@property Variant value() {
		return m_value;
	}

	@property Variant mergedValue() {
		if( m_merged ) return m_refToValue.value;
		else return m_value;
	}

	@property TypeInfo type() {
		return m_value.type;
	}
	
	unittest {
		int value = 10;
		Cell cell = new Cell(value);
		assert( cell.type is typeid(value) );
		assert( cell.value == value );
		Cell anotherCell = new Cell(11, cell);
		assert( anotherCell.mergedValue == value );
	}
}

unittest {
	mixin( unit!"xlsx_reader" );
	Workbook workbook = new Workbook("../../resource/excel_workbook.xlsx");
	Sheet sheet1 = workbook.sheetByName("Лист1");
	auto row0 = sheet1.row(0);
	assert( row0[0].value == "Cell-A1" );
	assert( sheet1.cell("A1").value == "Cell-A1" );
	assert( row0[1].value == "Cell-B1" );
	assert( row0[2].value == "Cell-C1" );
	assert( sheet1.cell("A4").value == 1 );
	assert( sheet1.cell("A5").value == 50.3 );
	writeln( sheet1.cell("A8").value );
}

int main(string[] argv) {
	return 0;
}