#include <XLSX_OpenXLSX_Document.h>
#include <XLSX_OpenXLSX_Workbook.h>

XLSX::OpenXLSX::Document::Document(const std::string& _Path)
{
    if(_Path.empty())
        return;

    try
    {
        m_Document.open(_Path);
    }
    catch(...)
    {
    }
}

XLSX::OpenXLSX::Document::~Document()
{
    if(m_Document.isOpen())
        m_Document.close();
}

bool XLSX::OpenXLSX::Document::open(const std::string& _Path)
{
    try
    {
        m_Document.open(_Path);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

void XLSX::OpenXLSX::Document::close()
{
    try
    {
        if(m_Document.isOpen())
            m_Document.close();
    }
    catch(...)
    {
    }
}

bool XLSX::OpenXLSX::Document::save()
{
    try
    {
        if(m_Document.isOpen())
            m_Document.save();
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::OpenXLSX::Document::saveAs(const std::string& _Path)
{
    try
    {
        if(m_Document.isOpen())
            m_Document.saveAs(_Path, true);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

XLSX::IWorkbook::WorkbookPointer XLSX::OpenXLSX::Document::WorkBook()
{
    return XLSX::IWorkbook::WorkbookPointer(new Workbook(m_Document));
}
