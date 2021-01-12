#include "stdafx.h"
#include "AcWordXml.h"


AcWordXml::AcWordXml()
{
	//创建一个XML的文档对象。
	myDocument = new TiXmlDocument();
	TiXmlDeclaration* decl = new TiXmlDeclaration("1.0", "GB2312", "");
	myDocument->LinkEndChild(decl);
	//创建一个根元素并连接。
	RootElement = new TiXmlElement("w:wordDocument");
	RootElement->SetAttribute("xmlns:w", "http://schemas.microsoft.com/office/word/2003/wordml");
	RootElement->SetAttribute("xmlns:wx", "http://schemas.microsoft.com/office/word/2003/auxHint");
	RootElement->SetAttribute("xml:space", "preserve");
	myDocument->LinkEndChild(RootElement);
	bodyElement = new TiXmlElement("w:body");
	RootElement->LinkEndChild(bodyElement);
}


AcWordXml::~AcWordXml()
{
	delete myDocument;
}


CString  adjustAlignment( BOOL Alignment)
{
	CString temp = "center";
	switch (Alignment)
	{
	case  TableAlignment::TableLeft:
			return "start";
			break;
	case  TableAlignment::Tableright:
		return  "end";
		break;
	case  ParagraphAlignment::Paragraphboth:
		return  "both";
		break;
	case ParagraphAlignment::ParagraphLeft:
		return  "left";
		break;
	case ParagraphAlignment::Paragraphright:
		return  "right";
		break;
	default:
		return  "center";
	}
}
void tableBorder(TiXmlElement *tbPosit, TiXmlElement *tbBd)
{
	tbPosit->SetAttribute("w:val", "single");
	tbPosit->SetAttribute("w:sz", "4");
	tbPosit->SetAttribute("w:space", "0");
	tbPosit->SetAttribute("w:color", "auto");
	tbBd->LinkEndChild(tbPosit);
}

void AcWordXml::CreateXmlTable(std::vector<std::vector<CString>>&textArray,const BOOL TableAlignment,const int tableTextsize, const BOOL tableBorden,const int totalTablewidth)
{
	if (textArray.size() <= 0) return;
	std::vector<CString>&lineArray = textArray[0];
	if (lineArray.size() <= 0)  return;

	long nRows = textArray.size();
	long nColumns = lineArray.size();
	TiXmlElement *table = new TiXmlElement("w:tbl");
	bodyElement->LinkEndChild(table);

	
	TiXmlElement *tblPr = new TiXmlElement("w:tblPr");
	table->LinkEndChild(tblPr);

	//adjust table alignment
	TiXmlElement *tbjc = new TiXmlElement("w:jc");
	CString temp = adjustAlignment(TableAlignment);
	tbjc->SetAttribute("w:val", temp);
	tblPr->LinkEndChild(tbjc);

	//adjust everygrid width ;there are some errors
	//TiXmlElement *tbGridWidth = new TiXmlElement("w:tblGrid");
	//table->LinkEndChild(tbGridWidth);
	//TiXmlElement *tblGridCol = new TiXmlElement("w:gridCol");
	//tblGridCol->SetAttribute("w:w", "2500");
	//tbGridWidth->LinkEndChild(tblGridCol);
	//TiXmlElement *tblGridCol2 = new TiXmlElement("w:gridCol");
	//tblGridCol2->SetAttribute("w:w", "3000");
	//tbGridWidth->LinkEndChild(tblGridCol2);

	//adjust total table width
	TiXmlElement *tblW = new TiXmlElement("w:tblW");
	tblPr->LinkEndChild(tblW);
	tblW->SetAttribute("w:type", "dxa");
	tblW->SetAttribute("w:w", totalTablewidth);

	
	TiXmlElement *tblLook= new TiXmlElement("w:tblLook");
	tblLook->SetAttribute("w:val", "04A0");
	tblPr->LinkEndChild(tblLook);

	//adjust border
	TiXmlElement *tbBd = new TiXmlElement("w:tblBorders");
	tblPr->LinkEndChild(tbBd);
	TiXmlElement *tbPosit;
	CString tempPos = "w:top";
	tbPosit = new TiXmlElement(tempPos);
	tableBorder(tbPosit, tbBd);
	TiXmlElement *tbPosit2;
	tempPos = "w:bottom";
	tbPosit2 = new TiXmlElement(tempPos);
	tableBorder(tbPosit2, tbBd);
	TiXmlElement *tbPosit3;
	tempPos = "w:left";
	tbPosit3 = new TiXmlElement(tempPos);
	tableBorder(tbPosit3, tbBd);
	TiXmlElement *tbPosit4;
	tempPos = "w:right";
	tbPosit4 = new TiXmlElement(tempPos);
	tableBorder(tbPosit4, tbBd);
	TiXmlElement *tbPosit5;
	tempPos = "w:insideH";
	tbPosit5 = new TiXmlElement(tempPos);
	tableBorder(tbPosit5, tbBd);
	TiXmlElement *tbPosit6;
	tempPos = "w:insideV";
	tbPosit6 = new TiXmlElement(tempPos);
	tableBorder(tbPosit6, tbBd);

	for (int i=0;i<nRows;i++)
	{
		TiXmlElement *tabler = new TiXmlElement("w:tr");
		table->LinkEndChild(tabler);
		std::vector<CString>&tempArray = textArray[i];
		for (int j=0;j<nColumns;j++)
		{
			TiXmlElement *tc = new TiXmlElement("w:tc");
			tabler->LinkEndChild(tc);
			TiXmlElement *tcPr = new TiXmlElement("w:tcPr");
			tabler->LinkEndChild(tcPr);
			TiXmlElement *tcW = new TiXmlElement("w:tcW");
			tcW->SetAttribute("w:w", "3500");
			tcW->SetAttribute("w:type","dxa");
			tcPr->LinkEndChild(tcW);
			TiXmlElement *tp2 = new TiXmlElement("w:p");
			tc->LinkEndChild(tp2);
			CreateXmlText(tp2, tempArray[j], tableBorden, tableTextsize);
		}
	}
}

 /*newline:文字上下空格
textspace：文字之间间隔
textClr:16进制颜色代码
*/
void AcWordXml::CreateXmlText(TiXmlElement *paragraph, const CString text, const bool bBorden , const int textsize ,const CString textClr, const bool bNewLine, const  int bEachTextSpace )
{	
	TiXmlElement *tr = new TiXmlElement("w:r");
	paragraph->LinkEndChild(tr);
	TiXmlElement *rPr = new TiXmlElement("w:rPr");
	tr->LinkEndChild(rPr);
	if (bBorden == true)
	{
		TiXmlElement *textb = new TiXmlElement("w:b");
		rPr->LinkEndChild(textb);
	}
	
	TiXmlElement *textcolor = new TiXmlElement("w:color");
	rPr->LinkEndChild(textcolor);
	textcolor->SetAttribute("w:val", textClr);

	TiXmlElement *texts= new TiXmlElement("w:sz");
	rPr->LinkEndChild(texts);
	texts->SetAttribute("w:val", textsize);
	TiXmlElement *tt = new TiXmlElement("w:t");
	if (bEachTextSpace !=0)
	{
		TiXmlElement *textspacing= new TiXmlElement("w:spacing");
		textspacing->SetAttribute("w:val", bEachTextSpace);
		rPr->LinkEndChild(textspacing);
		tt->SetAttribute("xml:space", "preserve");
	}
	if (bNewLine == true)
	{
		TiXmlElement *textbr = new TiXmlElement("w:br");
		tt->LinkEndChild(textbr);
	}
	tt->LinkEndChild(new TiXmlText(text));
	tr->LinkEndChild(tt);
}

void AcWordXml::CreateXmlaragraph(const BOOL paragraphAlignment,const int pbefore,const int pafter)
{
	TiXmlElement *paragraphNew = new TiXmlElement("w:p");
	paragraph = paragraphNew;
	bodyElement->LinkEndChild(paragraph);
	
	TiXmlElement *pPr = new TiXmlElement("w:pPr");
	paragraphNew->LinkEndChild(pPr);
	TiXmlElement *pjc = new TiXmlElement("w:jc");
	CString temp = adjustAlignment(paragraphAlignment);
	pjc->SetAttribute("w:val",temp);
	pPr->LinkEndChild(pjc);
	
	TiXmlElement *pspace = new TiXmlElement("w:spacing");
	pspace->SetAttribute("w:before", pbefore);
	pspace->SetAttribute("w:after", pafter);
	pPr->LinkEndChild(pspace);
}
