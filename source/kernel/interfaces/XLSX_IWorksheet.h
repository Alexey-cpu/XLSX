#ifndef XLSX_IWORKSHEET_H
#define XLSX_IWORKSHEET_H

// custom
#include <XLSX_ICellData.h>

// STL
#include <memory>
#include <string>

namespace XLSX
{
    template<typename ... T> class IWorksheet;

    typedef std::shared_ptr<IWorksheet<>> WorksheetPointer;

    template<typename ... T>
    class IWorksheet
    {
    public:

        // constructors
        IWorksheet(){}

        // virtual destructor
        virtual ~IWorksheet(){}

        // getters
        virtual std::string getName() const
        {
            return std::string();
        }

        virtual int getIndex() const
        {
            return 0;
        }

        template<typename __type>
        __type getData(int _Row, int _Column) const
        {
            const ICellData<__type>* getter =
                dynamic_cast<const ICellData<__type>*>(this);

            return getter != nullptr ? getter->getData(_Row, _Column) : __type();
        }

        // setters
        virtual bool setName(std::string)
        {
            return false;
        }

        virtual bool setIndex(int)
        {
            return false;
        }

        template<typename __type>
        bool setData(int _Row, int _Column, __type _Value)
        {
            ICellData<__type>* setter =
                dynamic_cast<ICellData<__type>*>(this);

            return setter != nullptr && setter->setData(_Row, _Column, _Value);
        }

        // interface
        virtual int rowCount()
        {
            return 0;
        }

        virtual int columnCount()
        {
            return 0;
        }

        virtual bool insertRows(int _From, int _Count)
        {
            return false;
        }

        virtual bool insertColumns(int _From, int _Count)
        {
            return false;
        }

        virtual bool removeRows(int _Row, int _Count)
        {
            return false;
        }

        virtual bool removeColumns(int _Column, int _Count)
        {
            return false;
        }

        virtual bool moveRows(int from, int count, int to)
        {
            return false;
        }

        virtual bool moveColumns(int _From, int _Count, int _To)
        {
            return false;
        }

        std::string verticalHeader(int _Index) const
        {
            return std::to_string(_Index + 1);
        }

        std::string horizontalHeader(int _Index) const
        {
            ++_Index;
            std::string name;

            while (_Index > 0)
            {
                _Index--;
                name = (char)('A' + _Index%26) + name;
                _Index /= 26;
            }

            return name;
        }

        // static API
        static WorksheetPointer empty()
        {
            return WorksheetPointer(new IWorksheet<>());
        }
    };

    template< typename _Worksheet, typename ... _Others>
    class IWorksheet<_Worksheet, _Others...> : public IWorksheet<_Others...>
    {
    public:

        // constructors
        IWorksheet(_Worksheet _Sheet) : m_Sheet(_Sheet){}

        // virtual destructor
        virtual ~IWorksheet(){}

    protected:

        _Worksheet m_Sheet;
    };
};

#endif // XLSX_IWORKSHEET_H
