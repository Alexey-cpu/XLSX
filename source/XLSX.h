#ifndef XLSX_H
#define XLSX_H

// STL
#include <string>
#include <memory>
#include <vector>

namespace XLSXCpp
{
    // ICell
    template<typename __type>
    class ICell
    {
    public:

        // constructors
        ICell(){}

        // virtual destructor
        virtual ~ICell(){}

        // getters
        virtual __type getData(int _Row, int _Column, ICell<__type>* _Data = nullptr) const = 0;

        // setters
        virtual bool setData(int _Row, int _Column, __type _Value, ICell<__type>* _Data = nullptr) = 0;
    };

    // IWorksheet
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
            const ICell<__type>* getter =
                dynamic_cast<const ICell<__type>*>(this);

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
            ICell<__type>* setter =
                dynamic_cast<ICell<__type>*>(this);

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

        // static API
        static std::unique_ptr<IWorksheet<>> empty()
        {
            return std::unique_ptr<IWorksheet<>>(new IWorksheet<>());
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

    // IWorkbook
    class IWorkbook
    {
    public:

        typedef std::shared_ptr<IWorksheet<>> WorksheetPointer;

        // constructors
        IWorkbook(){}

        // virtual destructor
        virtual ~IWorkbook(){}

        // interface
        virtual WorksheetPointer findSheet(int _Index) const = 0;
        virtual WorksheetPointer findSheet(std::string _Name) const = 0;
        virtual WorksheetPointer addSheet(std::string _Name) const = 0;
        virtual bool removeSheet(std::string) = 0;
        virtual std::vector<WorksheetPointer> Sheets() const = 0;
        virtual int sheetsCount() const = 0;
        virtual bool insertSheets(int _From, int _Count) = 0;
        virtual bool removeSheets(int _From, int _Count) = 0;
    };

    // IDocument
    class IDocument
    {
    public:

        typedef std::unique_ptr<IWorkbook> WorkbookPointer;

        // constructors
        IDocument(){}

        // virtual destructor
        virtual ~IDocument(){}

        // interface
        virtual bool open(const std::string& _Path) = 0;
        virtual void close() = 0;
        virtual bool save()  = 0;
        virtual bool saveAs(const std::string& _Path) = 0;

        // sheet functions
        virtual WorkbookPointer WorkBook() const = 0;
    };
};

#endif // XLSX_H
