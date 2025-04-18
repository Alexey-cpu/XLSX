#ifndef XLSX_OPENXLSX_WORKBOOK_H
#define XLSX_OPENXLSX_WORKBOOK_H

// custom
#include "XLSX_IWorkbook.h"

// OpenXLSX
#include <OpenXLSX.hpp>

namespace XLSX
{
    namespace OpenXLSX
    {
        class Workbook : public XLSX::IWorkbook
        {
        public:

            // constructors
            Workbook(const ::OpenXLSX::XLDocument& _Document);

            // virtual destructor
            virtual ~Workbook();

            // sheet functions
            virtual WorksheetPointer findSheet(int _Index) const override;
            virtual WorksheetPointer findSheet(std::string _Name) const override;
            virtual WorksheetPointer addSheet(std::string _Name) const override;
            virtual bool removeSheet(std::string _Name) override;
            virtual std::vector<XLSX::WorksheetPointer> Sheets() const override;
            virtual int sheetsCount() const override;
            virtual bool insertSheets(int _From, int _Count) override;
            virtual bool removeSheets(int _From, int _Count) override;

        protected:

            const ::OpenXLSX::XLDocument& m_Document;
        };
    }
};

#endif // XLSX_OPENXLSX_WORKBOOK_H
