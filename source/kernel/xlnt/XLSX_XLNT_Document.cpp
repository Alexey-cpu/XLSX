#include <XLSX_XLNT_Document.h>
#include <XLSX_XLNT_Workbook.h>

XLSX::XLNT::Document::Document(const std::string& _Path)
{
    if(_Path.empty())
        return;

    try
    {
        m_Path = _Path;
        m_Document.load(m_Path);
    }
    catch(...)
    {
    }
}

XLSX::XLNT::Document::~Document()
{
    m_Document.clear();
}

bool XLSX::XLNT::Document::open(const std::string& _Path)
{
    try
    {
        m_Document.load(_Path);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

void XLSX::XLNT::Document::close()
{
    try
    {
        m_Document.clear();
    }
    catch(...)
    {
    }
}

bool XLSX::XLNT::Document::save()
{
    try
    {
        m_Document.save(m_Path);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::XLNT::Document::saveAs(const std::string& _Path)
{
    try
    {
        m_Path = _Path;

        return save();
    }
    catch(...)
    {
        return false;
    }

    return true;
}

XLSX::IWorkbook::WorkbookPointer XLSX::XLNT::Document::WorkBook()
{
    return XLSX::IWorkbook::WorkbookPointer(new Workbook(m_Document));
}
