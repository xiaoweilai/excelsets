#include "qexcel.h"

#include <QAxObject>
#include <QFile>
#include <QStringList>
#include <QDebug>
#include <QDir>


QEXCEL::QEXCEL(QString xlsFilePath, QObject *parent)
{
    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;

    excel = new QAxObject("Excel.Application", parent);
    workBooks = excel->querySubObject("Workbooks");
    QFile file(xlsFilePath);
    if (file.exists())
    {
        workBooks->dynamicCall("Open(const QString&)", xlsFilePath);
        workBook = excel->querySubObject("ActiveWorkBook");
        sheets = workBook->querySubObject("WorkSheets");
    }
    else
    {
        if(CreateExcel(xlsFilePath))
        {
            workBooks->dynamicCall("Open(const QString&)", xlsFilePath);
            workBook = excel->querySubObject("ActiveWorkBook");
            sheets = workBook->querySubObject("WorkSheets");
        }
    }
}

QEXCEL::~QEXCEL()
{
    close();
}


/**
* @brief Create Excel File
* @param file [QString]  the name of the opened file
* @return 0:success   -1:failed
*/
bool QEXCEL::CreateExcel(QString file)
{
    QDir  dTemp;

    if(dTemp.exists(file))
    {
        qDebug()<<" QExcel::CreateExcel: exist file"<<file;
        return false;
    }

    qDebug()<<" QExcel::CreateExcel: succes";

    /**< create new excel sheet file.*/
    QAxObject * workSheet = excel->querySubObject("WorkBooks");
    workSheet->dynamicCall("Add");

    /**< save Excel.*/
    QAxObject * workExcel= excel->querySubObject("ActiveWorkBook");
    excel->setProperty("DisplayAlerts", 0);
    workExcel->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",file,56,QString(""),QString(""),false,false);
    excel->setProperty("DisplayAlerts", 1);
    workExcel->dynamicCall("Close (Boolean)", false);

    /**< exit Excel.*/
    //excel->dynamicCall("Quit (void)");

    return true;
}

void QEXCEL::close()
{
    excel->dynamicCall("Quit()");

    delete sheet;
    delete sheets;
    delete workBook;
    delete workBooks;
    delete excel;

    excel = 0;
    workBooks = 0;
    workBook = 0;
    sheets = 0;
    sheet = 0;
}

QAxObject *QEXCEL::getWorkBooks()
{
    return workBooks;
}

QAxObject *QEXCEL::getWorkBook()
{
    return workBook;
}

QAxObject *QEXCEL::getWorkSheets()
{
    return sheets;
}

QAxObject *QEXCEL::getWorkSheet()
{
    return sheet;
}

void QEXCEL::selectSheet(const QString& sheetName)
{
    sheet = sheets->querySubObject("Item(const QString&)", sheetName);
}

void QEXCEL::deleteSheet(const QString& sheetName)
{
    QAxObject * a = sheets->querySubObject("Item(const QString&)", sheetName);
    a->dynamicCall("delete");
}

void QEXCEL::deleteSheet(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    a->dynamicCall("delete");
}

void QEXCEL::selectSheet(int sheetIndex)
{
    sheet = sheets->querySubObject("Item(int)", sheetIndex);
}

void QEXCEL::setCellString(int row, int column, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QEXCEL::setCellFontBold(int row, int column, bool isBold)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Bold", isBold);
}

void QEXCEL::setCellFontSize(int row, int column, int size)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

void QEXCEL::mergeCells(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

void QEXCEL::mergeCells(int topLeftRow, int topLeftColumn, int bottomRightRow, int bottomRightColumn)
{
    QString cell;
    cell.append(QChar(topLeftColumn - 1 + 'A'));
    cell.append(QString::number(topLeftRow));
    cell.append(":");
    cell.append(QChar(bottomRightColumn - 1 + 'A'));
    cell.append(QString::number(bottomRightRow));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("VerticalAlignment", -4108);//xlCenter
    range->setProperty("WrapText", true);
    range->setProperty("MergeCells", true);
}

QVariant QEXCEL::getCellValue(int row, int column)
{
    QAxObject *range = sheet->querySubObject("Cells(int,int)", row, column);
    return range->property("Value");
}

void QEXCEL::save()
{
    workBook->dynamicCall("Save()");
}

int QEXCEL::getSheetsCount()
{
    return sheets->property("Count").toInt();
}

QString QEXCEL::getSheetName()
{
    return sheet->property("Name").toString();
}

QString QEXCEL::getSheetName(int sheetIndex)
{
    QAxObject * a = sheets->querySubObject("Item(int)", sheetIndex);
    return a->property("Name").toString();
}

void QEXCEL::getUsedRange(int *topLeftRow, int *topLeftColumn, int *bottomRightRow, int *bottomRightColumn)
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    *topLeftRow = usedRange->property("Row").toInt();
    *topLeftColumn = usedRange->property("Column").toInt();

    QAxObject *rows = usedRange->querySubObject("Rows");
    *bottomRightRow = *topLeftRow + rows->property("Count").toInt() - 1;

    QAxObject *columns = usedRange->querySubObject("Columns");
    *bottomRightColumn = *topLeftColumn + columns->property("Count").toInt() - 1;
}

void QEXCEL::setColumnWidth(int column, int width)
{
    QString columnName;
    columnName.append(QChar(column - 1 + 'A'));
    columnName.append(":");
    columnName.append(QChar(column - 1 + 'A'));

    QAxObject * col = sheet->querySubObject("Columns(const QString&)", columnName);
    col->setProperty("ColumnWidth", width);
}

void QEXCEL::setCellTextCenter(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("HorizontalAlignment", -4108);//xlCenter
}

void QEXCEL::setCellTextWrap(int row, int column, bool isWrap)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("WrapText", isWrap);
}

void QEXCEL::setAutoFitRow(int row)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject * rows = sheet->querySubObject("Rows(const QString &)", rowsName);
    rows->dynamicCall("AutoFit()");
}

void QEXCEL::insertSheet(QString sheetName)
{
    sheets->querySubObject("Add()");
    QAxObject * a = sheets->querySubObject("Item(int)", 1);
    a->setProperty("Name", sheetName);
}

void QEXCEL::mergeSerialSameCellsInAColumn(int column, int topRow)
{
    int a,b,c,rowsCount;
    getUsedRange(&a, &b, &rowsCount, &c);

    int aMergeStart = topRow, aMergeEnd = topRow + 1;

    QString value;
    while(aMergeEnd <= rowsCount)
    {
        value = getCellValue(aMergeStart, column).toString();
        while(value == getCellValue(aMergeEnd, column).toString())
        {
            clearCell(aMergeEnd, column);
            aMergeEnd++;
        }
        aMergeEnd--;
        mergeCells(aMergeStart, column, aMergeEnd, column);

        aMergeStart = aMergeEnd + 1;
        aMergeEnd = aMergeStart + 1;
    }
}

void QEXCEL::clearCell(int row, int column)
{
    QString cell;
    cell.append(QChar(column - 1 + 'A'));
    cell.append(QString::number(row));

    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

void QEXCEL::clearCell(const QString& cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("ClearContents()");
}

int QEXCEL::getUsedRowsCount()
{
    QAxObject *usedRange = sheet->querySubObject("UsedRange");
    int topRow = usedRange->property("Row").toInt();
    QAxObject *rows = usedRange->querySubObject("Rows");
    int bottomRow = topRow + rows->property("Count").toInt() - 1;
    return bottomRow;
}

void QEXCEL::setCellString(const QString& cell, const QString& value)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->dynamicCall("SetValue(const QString&)", value);
}

void QEXCEL::setCellFontSize(const QString &cell, int size)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Size", size);
}

void QEXCEL::setCellTextCenter(const QString &cell)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("HorizontalAlignment", -4108);//xlCenter
}

void QEXCEL::setCellFontBold(const QString &cell, bool isBold)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range = range->querySubObject("Font");
    range->setProperty("Bold", isBold);
}

void QEXCEL::setCellTextWrap(const QString &cell, bool isWrap)
{
    QAxObject *range = sheet->querySubObject("Range(const QString&)", cell);
    range->setProperty("WrapText", isWrap);
}

void QEXCEL::setRowHeight(int row, int height)
{
    QString rowsName;
    rowsName.append(QString::number(row));
    rowsName.append(":");
    rowsName.append(QString::number(row));

    QAxObject * r = sheet->querySubObject("Rows(const QString &)", rowsName);
    r->setProperty("RowHeight", height);
}
