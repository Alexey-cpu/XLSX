#ifndef XLSX_IWORKBOOK_H
#define XLSX_IWORKBOOK_H

// custom
#include <XLSX_IWorksheet.h>

// STL
#include <memory>
#include <vector>

namespace XLSX
{
    class IWorkbook
    {
    public:

        typedef std::shared_ptr<IWorkbook> WorkbookPointer;

        // constructors
        IWorkbook(){}

        // virtual destructor
        virtual ~IWorkbook(){}

        // interface
        virtual IWorksheet<>::WorksheetPointer findSheet(int _Index) const = 0;
        virtual IWorksheet<>::WorksheetPointer findSheet(std::string _Name) const = 0;
        virtual IWorksheet<>::WorksheetPointer addSheet(std::string _Name) const = 0;
        virtual bool removeSheet(std::string) = 0;
        virtual std::vector<IWorksheet<>::WorksheetPointer> Sheets() const = 0;
        virtual int sheetsCount() const = 0;
        virtual bool insertSheets(int _From, int _Count) = 0;
        virtual bool removeSheets(int _From, int _Count) = 0;
    };
};

#endif // XLSX_IWORKBOOK_H
