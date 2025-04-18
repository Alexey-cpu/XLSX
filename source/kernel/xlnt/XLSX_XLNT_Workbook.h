#ifndef XLSX_XLNT_WORKBOOK_H
#define XLSX_XLNT_WORKBOOK_H

// custom
#include "XLSX_IWorkbook.h"

// OpenXLSX
#include <xlnt/xlnt.hpp>

namespace XLSX
{
    namespace XLNT
    {
        class Workbook : public XLSX::IWorkbook
        {
        public:

            // constructors
            Workbook(xlnt::workbook& _Document);

            // virtual destructor
            virtual ~Workbook();

            // sheet functions
            virtual XLSX::WorksheetPointer findSheet(int _Index) const override;
            virtual XLSX::WorksheetPointer findSheet(std::string _Name) const override;
            virtual XLSX::WorksheetPointer addSheet(std::string _Name) const override;
            virtual bool removeSheet(std::string _Name) override;
            virtual std::vector<XLSX::WorksheetPointer> Sheets() const override;
            virtual int sheetsCount() const override;
            virtual bool insertSheets(int _From, int _Count) override;
            virtual bool removeSheets(int _From, int _Count) override;

        protected:

            xlnt::workbook& m_Document;
        };
    }
};

#endif // XLSX_XLNT_WORKBOOK_H
