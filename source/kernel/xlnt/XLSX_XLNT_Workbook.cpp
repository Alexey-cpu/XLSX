#include <XLSX_XLNT_Workbook.h>
#include <XLSX_XLNT_Worksheet.h>

// STL
#include <string>

XLSX::XLNT::Workbook::Workbook(xlnt::workbook& _Document) :
    m_Document(_Document){}

XLSX::XLNT::Workbook::~Workbook(){}

XLSX::WorksheetPointer XLSX::XLNT::Workbook::findSheet(int _Index) const
{
    try
    {
        auto sheet = m_Document.sheet_by_index(_Index);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

XLSX::WorksheetPointer XLSX::XLNT::Workbook::findSheet(std::string _Name) const
{
    try
    {
        auto sheet = m_Document.sheet_by_title(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

XLSX::WorksheetPointer XLSX::XLNT::Workbook::addSheet(std::string _Name) const
{
    // retrieve existing sheet
    try
    {
        auto sheet = m_Document.sheet_by_title(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch(...)
    {
    }

    // create new sheet
    try
    {
        auto sheet = m_Document.create_sheet();
        sheet.title(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

bool XLSX::XLNT::Workbook::removeSheet(std::string _Name)
{
    try
    {
        auto sheet = m_Document.sheet_by_title(_Name);
        m_Document.remove_sheet(sheet);
        return true;
    }
    catch (...)
    {
        return false;
    }
}

std::vector<XLSX::WorksheetPointer> XLSX::XLNT::Workbook::Sheets() const
{
    std::vector<XLSX::WorksheetPointer> sheets;

    for(int i = 0; i < m_Document.sheet_count(); i++)
        sheets.push_back(findSheet(i));

    return sheets;
}

int XLSX::XLNT::Workbook::sheetsCount() const
{
    return m_Document.sheet_count();
}

bool XLSX::XLNT::Workbook::insertSheets(int _From, int _Count)
{
    for(int i = 0; i < _Count; i++)
        m_Document.create_sheet(i);

    return true;
}

bool XLSX::XLNT::Workbook::removeSheets(int _From, int _Count)
{
    auto sheets = Sheets();

    for(size_t i = _From; i < std::min(_From + _Count, sheetsCount()); i++)
        removeSheet(findSheet(i)->getName());

    return true;
}
