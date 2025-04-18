#ifndef XLSX_IDOCUMENT_H
#define XLSX_IDOCUMENT_H

#include <XLSX_IWorkbook.h>

namespace XLSX
{
    class IDocument
    {
    public:

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
        virtual WorkbookPointer WorkBook() = 0;
    };
};

#endif // XLSX_IDOCUMENT_H
