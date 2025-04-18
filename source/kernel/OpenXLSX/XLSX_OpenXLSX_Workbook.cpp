#include <XLSX_OpenXLSX_Workbook.h>
#include <XLSX_OpenXLSX_Workheet.h>

// STL
#include <set>
#include <string>

// XLSX::OpenXLSX::Workbook
XLSX::OpenXLSX::Workbook::Workbook(const ::OpenXLSX::XLDocument& _Document) :
    m_Document(_Document){}

XLSX::OpenXLSX::Workbook::~Workbook(){}

XLSX::WorksheetPointer XLSX::OpenXLSX::Workbook::findSheet(int _Index) const
{
    if(!m_Document.isOpen())
        return XLSX::IWorksheet<>::empty();

    try
    {
        auto sheet = m_Document.workbook().sheet(_Index + 1);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

XLSX::WorksheetPointer XLSX::OpenXLSX::Workbook::findSheet(std::string _Name) const
{
    if(!m_Document.isOpen())
        return XLSX::IWorksheet<>::empty();

    try
    {
        auto sheet = m_Document.workbook().sheet(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

XLSX::WorksheetPointer XLSX::OpenXLSX::Workbook::addSheet(std::string _Name) const
{
    if(!m_Document.isOpen())
        return XLSX::IWorksheet<>::empty();

    // retrieve existing sheet
    try
    {
        auto sheet = m_Document.workbook().worksheet(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch(...)
    {
    }

    // create new sheet
    try
    {
        m_Document.workbook().addWorksheet(_Name);
        auto sheet = m_Document.workbook().worksheet(_Name);
        return XLSX::WorksheetPointer(new Worksheet(sheet));
    }
    catch (...)
    {
        return XLSX::IWorksheet<>::empty();
    }
}

bool XLSX::OpenXLSX::Workbook::removeSheet(std::string _Name)
{
    if(!m_Document.isOpen())
        return false;

    try
    {
        m_Document.workbook().deleteSheet(_Name);
        return true;
    }
    catch (...)
    {
        return false;
    }
}

std::vector<XLSX::WorksheetPointer> XLSX::OpenXLSX::Workbook::Sheets() const
{
    if(!m_Document.isOpen())
        return std::vector<XLSX::WorksheetPointer>();

    std::vector<XLSX::WorksheetPointer> sheets;

    for(int i = 0; i < m_Document.workbook().sheetCount(); i++)
        sheets.push_back(findSheet(i));

    return sheets;
}

int XLSX::OpenXLSX::Workbook::sheetsCount() const
{
    return m_Document.isOpen() ? m_Document.workbook().sheetCount() : 0;
}

bool XLSX::OpenXLSX::Workbook::insertSheets(int _From, int _Count)
{
    if(!m_Document.isOpen())
        return false;

    std::set<std::string> seed;
    std::map<int, std::string> newNames;
    auto sheets = Sheets();

    for(auto& sheet : sheets)
        seed.insert(sheet->getName());

    for(int i = 0; i < _Count; i++)
    {
        int postfix = sheetsCount();
        std::string name = "List-" + std::to_string(postfix);

        while (seed.contains(name))
            name = "List-" + std::to_string(++postfix);

        std::cout << name << "\n";

        newNames.insert({_From + i, name});
        seed.insert(name);
    }

    for(auto& name : newNames)
    {
        auto sheet = addSheet(name.second);
        sheet->setIndex(name.first);
    }

    return true;
}

bool XLSX::OpenXLSX::Workbook::removeSheets(int _From, int _Count)
{
    if(!m_Document.isOpen())
        return false;

    auto sheets = Sheets();

    for(size_t i = _From; i < std::min(_From + _Count, sheetsCount()); i++)
        removeSheet(findSheet(i)->getName());

    return true;
}
