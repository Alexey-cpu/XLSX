#include <XLSX_XLNT_Worksheet.h>

XLSX::XLNT::Worksheet::Worksheet(const xlnt::worksheet& _Sheet) :
    XLSX::IWorksheet<xlnt::worksheet>(_Sheet){}

XLSX::XLNT::Worksheet::~Worksheet(){}

std::string XLSX::XLNT::Worksheet::getName() const
{
    return m_Sheet.title();
}

int XLSX::XLNT::Worksheet::getIndex() const
{
    return m_Sheet.workbook().index(m_Sheet);
}

std::string XLSX::XLNT::Worksheet::getData(int _Row, int _Column, ICellData<std::string>* _Data) const
{
    typedef std::string __type;

    try
    {
        return m_Sheet.cell(_Column + 1, _Row + 1).to_string();
    }
    catch(...)
    {
        return __type();
    }
}

double XLSX::XLNT::Worksheet::getData(int _Row, int _Column, ICellData<double>* _Data) const
{
    typedef double __type;

    try
    {
        return m_Sheet.cell(_Column + 1, _Row + 1).value<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

float XLSX::XLNT::Worksheet::getData(int _Row, int _Column, ICellData<float>* _Data) const
{
    typedef float __type;

    try
    {
        return m_Sheet.cell(_Column + 1, _Row + 1).value<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

int XLSX::XLNT::Worksheet::getData(int _Row, int _Column, ICellData<int>* _Data) const
{
    typedef int __type;

    try
    {
        return m_Sheet.cell(_Column + 1, _Row + 1).value<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

bool XLSX::XLNT::Worksheet::setName(std::string _Name)
{
    try
    {
        m_Sheet.title(_Name);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::XLNT::Worksheet::setIndex(int _Index)
{
    try
    {
        m_Sheet.workbook().index(m_Sheet, _Index);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::XLNT::Worksheet::setData(int _Row, int _Column, std::string _Value, ICellData<std::string>* _Data)
{
    typedef std::string __type;

    try
    {
        m_Sheet.cell(_Column + 1, _Row + 1).value(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::XLNT::Worksheet::setData(int _Row, int _Column, double _Value, ICellData<double>* _Data)
{
    typedef double __type;

    try
    {
        m_Sheet.cell(_Column + 1, _Row + 1).value(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::XLNT::Worksheet::setData(int _Row, int _Column, float _Value, ICellData<float>* _Data)
{
    typedef float __type;

    try
    {
        m_Sheet.cell(_Column + 1, _Row + 1).value(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::XLNT::Worksheet::setData(int _Row, int _Column, int _Value, ICellData<int>* _Data)
{
    typedef int __type;

    try
    {
        m_Sheet.cell(_Column + 1, _Row + 1).value(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

int XLSX::XLNT::Worksheet::rowCount()
{
    return m_Sheet.rows(false).length();
}

int XLSX::XLNT::Worksheet::columnCount()
{
    return m_Sheet.columns(false).length();
}

bool XLSX::XLNT::Worksheet::insertRows(int _From, int _Count)
{
    if(m_Sheet.columns().length() <= 0)
    {
        try
        {
            m_Sheet.insert_columns(1, 1);

            if(_Count <= 1)
                return true;
        }
        catch(...)
        {
            return false;
        }
    }

    try
    {
        m_Sheet.insert_rows(_From + 1, _Count);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::XLNT::Worksheet::insertColumns(int _From, int _Count)
{
    if(rowCount() <= 0)
    {
        try
        {
            m_Sheet.insert_rows(1, _Count);

            if(_Count <= 1)
                return true;
        }
        catch(...)
        {
            return false;
        }
    }

    try
    {
        m_Sheet.insert_columns(_From + 1, _Count);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::XLNT::Worksheet::removeRows(int _From, int _Count)
{
    if(_From >= rowCount())
        return false;

    try
    {
        m_Sheet.delete_rows(_From + 1, _Count);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::XLNT::Worksheet::removeColumns(int _From, int _Count)
{
    if(_From >= columnCount())
        return false;

    try
    {
        m_Sheet.delete_columns(_From + 1, _Count);
    }
    catch(...)
    {
        return false;
    }

    return true;
}

bool XLSX::XLNT::Worksheet::moveRows(int _From, int _Count, int _To)
{
    if(rowCount()    <= 0   ||
        columnCount() <= 0  ||
        _From >= rowCount() ||
        _To >=  rowCount()  ||
        _From < 0           ||
        _To < 0             ||
        _From == _To)
        return false;

    if(_From > _To)
        std::swap(_From, _To);

    // retrieve source values
    for(int i = 0; i < _Count; i++)
    {
        int source = (_From + i) % std::max(rowCount(), 1) + 1;
        int target = (_To + i) % std::max(rowCount(), 1) + 1;

        for(int j = 0; j < columnCount(); j++)
        {
            auto sc = m_Sheet.cell(j + 1, source).to_string();
            auto tc = m_Sheet.cell(j + 1, target).to_string();

            // move source
            try
            {
                m_Sheet.cell(j + 1, source).value(std::stod(tc));
            }
            catch(...)
            {
                m_Sheet.cell(j + 1, source).value(tc);
            }

            // move target
            try
            {
                m_Sheet.cell(j + 1, target).value(std::stod(sc));
            }
            catch(...)
            {
                m_Sheet.cell(j + 1, target).value(sc);
            }
        }
    }

    return true;
}

bool XLSX::XLNT::Worksheet::moveColumns(int _From, int _Count, int _To)
{
    if(rowCount()    <= 0                   ||
        columnCount() <= 0                  ||
        _From >= columnCount()              ||
        _To >=  columnCount()               ||
        _From + _Count - 1 >= columnCount() ||
        _To + _Count - 1 >=  columnCount()  ||
        _From < 0                           ||
        _To < 0)
        return false;

    if(_From > _To)
        std::swap(_From, _To);

    for(int j = 0; j < _Count; j++)
    {
        int source = (_From + j) % std::max(columnCount(), 1) + 1;
        int target = (_To + j) % std::max(columnCount(), 1) + 1;

        // swap
        for(int k = 0 ; k < rowCount(); k++)
        {
            auto sc = m_Sheet.cell(source, k + 1).to_string();
            auto tc = m_Sheet.cell(target, k + 1).to_string();

            // move source
            try
            {
                m_Sheet.cell(source, k + 1).value(std::stod(tc));
            }
            catch(...)
            {
                m_Sheet.cell(source, k + 1).value(tc);
            }

            // move target
            try
            {
                m_Sheet.cell(target, k + 1).value(std::stod(sc));
            }
            catch(...)
            {
                m_Sheet.cell(target, k + 1).value(sc);
            }
        }
    }

    return true;
}
