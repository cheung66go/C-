#pragma once
#include "tinyxml.h"
#include"tinystr.h"
#include "vector"
enum TableAlignment
{
	TableCenter = 0,
	TableLeft = 1,
	Tableright = 2,
};
enum ParagraphAlignment
{
	Paragraphboth =4,
	ParagraphLeft =5,
	Paragraphright =6,
};
class AcWordXml
{
public:
	AcWordXml();
	~AcWordXml();
public:
		void CreateXmlTable(std::vector<std::vector<CString>>& textArray, const BOOL TableAlignment = TableCenter);
		void CreateXmlText(TiXmlElement *paragraph , const CString text, const bool bBorden = false, const int textsize = 20 , const CString textClr="auto", const bool bNewLine=false, const  int bEachTextSpace = 0);
		void CreateXmlaragraph(const BOOL paragraphAlignment=Paragraphboth);
public:

public:
	TiXmlDocument *myDocument;
	TiXmlElement *RootElement;
	TiXmlElement *bodyElement;
	TiXmlElement *paragraph;
};

