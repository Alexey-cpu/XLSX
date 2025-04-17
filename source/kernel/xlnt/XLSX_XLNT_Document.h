#ifndef XLSX_XLNT_DOCUMENT_H
#define XLSX_XLNT_DOCUMENT_H

// custom
#include "XLSX_IDocument.h"

// OpenXLSX
#include <xlnt/xlnt.hpp>

// STL
#include <filesystem>

namespace XLSX
{
    namespace XLNT
    {
        class Document : public XLSX::IDocument
        {
        public:

            // constructors
            Document(const std::string& _Path = std::string());

            // virtual destructor
            virtual ~Document();

            // interface
            virtual bool open(const std::string& _Path) override;
            virtual void close() override;
            virtual bool save() override;
            virtual bool saveAs(const std::string& _Path) override;
            virtual XLSX::IWorkbook::WorkbookPointer WorkBook() override;

        protected:

            xlnt::workbook m_Document;
            std::filesystem::path m_Path;
        };
    }
};

#endif // XLSX_XLNT_DOCUMENT_H
