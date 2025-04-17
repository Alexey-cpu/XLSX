#include <XLSX_OpenXLSX_Workheet.h>

XLSX::OpenXLSX::Worksheet::Worksheet(const ::OpenXLSX::XLWorksheet& _Sheet) :
    XLSX::IWorksheet<::OpenXLSX::XLWorksheet>(_Sheet){}

XLSX::OpenXLSX::Worksheet::~Worksheet(){}

std::string XLSX::OpenXLSX::Worksheet::getName() const
{
    return m_Sheet.name();
}

int XLSX::OpenXLSX::Worksheet::getIndex() const
{
    return m_Sheet.index();
}

std::string XLSX::OpenXLSX::Worksheet::getData(int _Row, int _Column, ICell<std::string>* _Data) const
{
    typedef std::string __type;

    try
    {
        return m_Sheet.cell(_Row + 1, _Column + 1).getString();
    }
    catch(...)
    {
        return __type();
    }
}

double XLSX::OpenXLSX::Worksheet::getData(int _Row, int _Column, ICell<double>* _Data) const
{
    typedef double __type;

    try
    {
        return m_Sheet.cell(_Row + 1, _Column + 1).value().get<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

float XLSX::OpenXLSX::Worksheet::getData(int _Row, int _Column, ICell<float>* _Data) const
{
    typedef float __type;

    try
    {
        return m_Sheet.cell(_Row + 1, _Column + 1).value().get<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

int XLSX::OpenXLSX::Worksheet::getData(int _Row, int _Column, ICell<int>* _Data) const
{
    typedef int __type;

    try
    {
        return m_Sheet.cell(_Row + 1, _Column + 1).value().get<__type>();
    }
    catch(...)
    {
        return __type();
    }
}

bool XLSX::OpenXLSX::Worksheet::setName(std::string _Name)
{
    try
    {
        m_Sheet.setName(_Name);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::OpenXLSX::Worksheet::setIndex(int _Index)
{
    try
    {
        m_Sheet.setIndex(_Index + 1);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::OpenXLSX::Worksheet::setData(int _Row, int _Column, std::string _Value, ICell<std::string>* _Data)
{
    typedef std::string __type;

    try
    {
        m_Sheet.cell(_Row + 1, _Column + 1).value().set<__type>(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::OpenXLSX::Worksheet::setData(int _Row, int _Column, double _Value, ICell<double>* _Data)
{
    typedef double __type;

    try
    {
        m_Sheet.cell(_Row + 1, _Column + 1).value().set<__type>(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::OpenXLSX::Worksheet::setData(int _Row, int _Column, float _Value, ICell<float>* _Data)
{
    typedef float __type;

    try
    {
        m_Sheet.cell(_Row + 1, _Column + 1).value().set<__type>(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

bool XLSX::OpenXLSX::Worksheet::setData(int _Row, int _Column, int _Value, ICell<int>* _Data)
{
    typedef int __type;

    try
    {
        m_Sheet.cell(_Row + 1, _Column + 1).value().set<__type>(_Value);

        return true;
    }
    catch(...)
    {
        return false;
    }
}

int XLSX::OpenXLSX::Worksheet::rowCount()
{
    return m_Sheet.rowCount();
}

int XLSX::OpenXLSX::Worksheet::columnCount()
{
    return m_Sheet.columnCount();
}

bool XLSX::OpenXLSX::Worksheet::insertRows(int _From, int _Count)
{
    (void)_Count;
    (void)_From;

    // append rows at the end
    for(int i = 0, j = std::max<int>(m_Sheet.rowCount() + 1, 1); i < _Count ; i++, j++)
        m_Sheet.row(j).values();

    return moveRows(rowCount() - _Count, _Count, _From + 1);
}

bool XLSX::OpenXLSX::Worksheet::insertColumns(int _From, int _Count)
{
    (void)_Count;
    (void)_From;

    int start = m_Sheet.columnCount();
    int end   = start + _Count;

    for(int j = start; j < end; j++)
        m_Sheet.cell(1, j+1).value() = std::string();

    return moveColumns(columnCount() - _Count, _Count, _From + 1);
}

bool XLSX::OpenXLSX::Worksheet::removeRows(int _From, int _Count)
{
    if(_From >= rowCount())
        return false;

    for(int i = 0; i < _Count; i++)
    {
        int index = _From + i;

        if(index < rowCount())
            m_Sheet.deleteRow(index + 1); // it deletes rows data but not rows !!!
    }

    return moveRows(rowCount() - _Count, _Count, _From + 1);
}

bool XLSX::OpenXLSX::Worksheet::removeColumns(int _From, int _Count)
{
    if(_From >= columnCount())
        return false;

    for(int i = 0; i < rowCount(); i++)
    {
        for(int j = 0; j < _Count; j++)
        {
            int index = _From + j;

            if(index < columnCount())
                m_Sheet.cell(i + 1, index + 1).value() = ::OpenXLSX::XLCellValue();
        }
    }

    return moveColumns(columnCount() - _Count, _Count, _From + 1);
}

bool XLSX::OpenXLSX::Worksheet::moveRows(int _From, int _Count, int _To)
{
    if(rowCount()    <= 0   ||
        columnCount() <= 0  ||
        _From >= rowCount() ||
        _To >=  rowCount()  ||
        _From < 0           ||
        _To < 0             ||
        _From == _To)
        return false;

    // retrieve source values
    if(_From > _To)
    {
        for(int i = _From; i < _From + _Count; i++)
        {
            for(int j = i - 1; j >= _To; j--)
            {
                int source = j;
                int target = j + 1;

                std::vector<::OpenXLSX::XLCellValue> a = m_Sheet.row(source+1).values();
                std::vector<::OpenXLSX::XLCellValue> b = m_Sheet.row(target+1).values();

                if(b.empty())
                    m_Sheet.row(source+1).values().clear();
                else
                    m_Sheet.row(source+1).values() = b;

                if(a.empty())
                    m_Sheet.row(target+1).values().clear();
                else
                    m_Sheet.row(target+1).values() = a;
            }
        }
    }
    else
    {
        for(int i = _From; i < _From + _Count; i++)
        {
            for(int j = i + 1; j < _To; j++)
            {
                int source = j;
                int target = j - 1;

                std::vector<::OpenXLSX::XLCellValue> a = m_Sheet.row(source+1).values();
                std::vector<::OpenXLSX::XLCellValue> b = m_Sheet.row(target+1).values();

                if(b.empty())
                    m_Sheet.row(source+1).values().clear();
                else
                    m_Sheet.row(source+1).values() = b;

                if(a.empty())
                    m_Sheet.row(target+1).values().clear();
                else
                    m_Sheet.row(target+1).values() = a;
            }
        }
    }

    return true;
}

bool XLSX::OpenXLSX::Worksheet::moveColumns(int _From, int _Count, int _To)
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
    {
        for(int i = _From; i < _From + _Count; i++)
        {
            for(int j = i - 1; j >= _To; j--)
            {
                int source = j;
                int target = j + 1;

                // swap
                for(int k = 0 ; k < m_Sheet.rowCount() ; k++)
                {
                    auto a = m_Sheet.cell(k+1, source+1).value().getString();
                    auto b = m_Sheet.cell(k+1, target+1).value().getString();

                    m_Sheet.cell(k+1, source+1).value() = b;
                    m_Sheet.cell(k+1, target+1).value() = a;
                }
            }
        }
    }
    else
    {
        for(int i = _From; i < _From + _Count; i++)
        {
            for(int j = i + 1; j < _To; j++)
            {
                int source = j;
                int target = j - 1;

                // swap
                for(int k = 0 ; k < m_Sheet.rowCount() ; k++)
                {
                    auto a = m_Sheet.cell(k+1, source+1).value().getString();
                    auto b = m_Sheet.cell(k+1, target+1).value().getString();

                    m_Sheet.cell(k+1, source+1).value() = b;
                    m_Sheet.cell(k+1, target+1).value() = a;
                }
            }
        }
    }

    return true;
}
