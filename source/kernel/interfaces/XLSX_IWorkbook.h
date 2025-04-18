#ifndef XLSX_IWORKBOOK_H
#define XLSX_IWORKBOOK_H

// custom
#include <XLSX_IWorksheet.h>

// STL
#include <memory>
#include <vector>

namespace XLSX
{
    class IWorkbook;

    typedef std::shared_ptr<IWorkbook> WorkbookPointer;

    class IWorkbook
    {
    public:

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

        virtual bool moveSheets(int _From, int _Count, int _To)
        {
            for(size_t i = 0; i < _Count; i++)
            {
                int source = (_From + i) ;
                int target = (_To + i);

                findSheet(source)->setIndex(target);
            }

            return true;
        }
    };
};

#endif // XLSX_IWORKBOOK_H
