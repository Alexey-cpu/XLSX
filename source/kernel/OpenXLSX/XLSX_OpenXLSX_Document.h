#ifndef XLSX_OPENXLSX_DOCUMENT_H
#define XLSX_OPENXLSX_DOCUMENT_H

// custom
#include "XLSX_IDocument.h"

// OpenXLSX
#include <OpenXLSX.hpp>

// XLSX::OpenXLSX
namespace XLSX
{
    // OpenXLSX wrapper
    namespace OpenXLSX
    {
        // Document
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

            ::OpenXLSX::XLDocument m_Document;
        };
    }
};

#endif // XLSX_OPENXLSX_DOCUMENT_H
