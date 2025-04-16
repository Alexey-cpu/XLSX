#ifndef XLSX_OPENXLSX_H
#define XLSX_OPENXLSX_H

// custom
#include "XLSX.h"

// OpenXLSX
#include <OpenXLSX.hpp>

// STL
#include <set>

// XLSXCpp::OpenXLSX
namespace XLSXCpp
{
    // OpenXLSX wrapper
    namespace OpenXLSX
    {
        // Worksheet
        class Worksheet :
            public XLSXCpp::IWorksheet<::OpenXLSX::XLWorksheet>,
            public XLSXCpp::ICell<std::string>,
            public XLSXCpp::ICell<double>,
            public XLSXCpp::ICell<float>,
            public XLSXCpp::ICell<int>
        {
        public:

            Worksheet(const ::OpenXLSX::XLWorksheet& _Sheet) :
                XLSXCpp::IWorksheet<::OpenXLSX::XLWorksheet>(_Sheet){}

            virtual ~Worksheet(){}

            // getters
            virtual std::string getName() const override
            {
                return m_Sheet.name();
            }

            virtual int getIndex() const override
            {
                return m_Sheet.index();
            }

            virtual std::string getData(int _Row, int _Column, ICell<std::string>* _Data = nullptr) const override
            {
                typedef std::string __type;

                try
                {
                    return m_Sheet.cell(_Row + 1, _Column + 1).value().get<__type>();
                }
                catch(...)
                {
                    return __type();
                }
            }

            virtual double getData(int _Row, int _Column, ICell<double>* _Data = nullptr) const override
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

            virtual float getData(int _Row, int _Column, ICell<float>* _Data = nullptr) const override
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

            virtual int getData(int _Row, int _Column, ICell<int>* _Data = nullptr) const override
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

            // setters
            virtual bool setName(std::string _Name) override
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

            virtual bool setIndex(int _Index) override
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

            virtual bool setData(int _Row, int _Column, std::string _Value, ICell<std::string>* _Data = nullptr) override
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

            virtual bool setData(int _Row, int _Column, double _Value, ICell<double>* _Data = nullptr) override
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

            virtual bool setData(int _Row, int _Column, float _Value, ICell<float>* _Data = nullptr) override
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

            virtual bool setData(int _Row, int _Column, int _Value, ICell<int>* _Data = nullptr) override
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

            // interface
            virtual int rowCount() override
            {
                return m_Sheet.rowCount();
            }

            virtual int columnCount() override
            {
                return m_Sheet.columnCount();
            }

            virtual bool insertRows(int _From, int _Count) override
            {
                // tweak starting row
                if(_From > rowCount())
                    _From = rowCount();

                // append rows at the end
                for(int i = 0, j = std::max<int>(m_Sheet.rowCount() + 1, 1); i < _Count ; i++, j++)
                    m_Sheet.row(j).values();

                // move added rows to an appropariate index
                return moveRows(rowCount() - _Count, _Count, _From);
            }

            virtual bool insertColumns(int _From, int _Count) override
            {
                // tweak starting column
                if(_From > columnCount())
                    _From = columnCount();

                int start = m_Sheet.columnCount();
                int end   = start + _Count;

                for(int j = start; j < end; j++)
                    m_Sheet.cell(1, j+1).value() = std::string();

                // move added columjns to an appropariate index
                return moveColumns(columnCount() - _Count, _Count, _From);
            }

            virtual bool removeRows(int _From, int _Count) override
            {
                if(_From >= rowCount())
                    return false;

                for(int i = 0; i < _Count; i++)
                {
                    int index = _From + i;

                    if(index < rowCount())
                        m_Sheet.deleteRow(index + 1); // it deletes rows data but not rows !!!
                }

                return moveRows(_From, _Count, rowCount() - _Count);
            }

            virtual bool removeColumns(int _From, int _Count) override
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

                return moveColumns(_From, _Count, columnCount() - _Count);
            }

            virtual bool moveRows(int _From, int _Count, int _To) override
            {
                if(rowCount()    <= 0   ||
                    columnCount() <= 0  ||
                    _From >= rowCount() ||
                    _To >=  rowCount()  ||
                    _From < 0           ||
                    _To < 0             ||
                    _From == _To)
                    return false;

                if(_From > _To)
                    std::swap(_From, _To);

                // retrieve source values
                for(int i = 0; i < _Count; i++)
                {
                    for(int j = 0; j < columnCount(); j++)
                    {
                        int source = (_From + i) % std::max(rowCount(), 1);
                        int target = (_To + i) % std::max(rowCount(), 1);

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

                return true;
            }

            virtual bool moveColumns(int _From, int _Count, int _To) override
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
                    std::swap(_From, _To);

                for(int i = 0; i < rowCount(); i++)
                {
                    for(int j = 0; j < _Count; j++)
                    {
                        int source = (_From + j) % std::max(columnCount(), 1);
                        int target = (_To + j) % std::max(columnCount(), 1);

                        // swap
                        for(int k = 0 ; k < m_Sheet.rowCount() ; k++)
                        {
                            auto a = (::OpenXLSX::XLCellValue)m_Sheet.cell(k+1, source+1).value();
                            auto b = (::OpenXLSX::XLCellValue)m_Sheet.cell(k+1, target+1).value();
                            m_Sheet.cell(k+1, source+1).value() = b;
                            m_Sheet.cell(k+1, target+1).value() = a;
                        }
                    }
                }

                return true;
            }
        };

        // Workbook
        class Workbook : public XLSXCpp::IWorkbook
        {
        public:

            // constructors
            Workbook(const ::OpenXLSX::XLDocument& _Document) :
                m_Document(_Document){}

            // virtual destructor
            virtual ~Workbook(){}

            // sheet functions
            virtual WorksheetPointer findSheet(int _Index) const override
            {
                if(!m_Document.isOpen())
                    return XLSXCpp::IWorksheet<>::empty();

                try
                {
                    auto sheet = m_Document.workbook().sheet(_Index + 1);
                    return WorksheetPointer(new Worksheet(sheet));
                }
                catch (...)
                {
                    return XLSXCpp::IWorksheet<>::empty();
                }
            }

            virtual WorksheetPointer findSheet(std::string _Name) const override
            {
                if(!m_Document.isOpen())
                    return XLSXCpp::IWorksheet<>::empty();

                try
                {
                    auto sheet = m_Document.workbook().sheet(_Name);
                    return WorksheetPointer(new Worksheet(sheet));
                }
                catch (...)
                {
                    return XLSXCpp::IWorksheet<>::empty();
                }
            }

            virtual WorksheetPointer addSheet(std::string _Name) const override
            {
                if(!m_Document.isOpen())
                    return XLSXCpp::IWorksheet<>::empty();

                // retrieve existing sheet
                try
                {
                    auto sheet = m_Document.workbook().worksheet(_Name);
                    return WorksheetPointer(new Worksheet(sheet));
                }
                catch(...)
                {
                }

                // create new sheet
                try
                {
                    m_Document.workbook().addWorksheet(_Name);
                    auto sheet = m_Document.workbook().worksheet(_Name);
                    return WorksheetPointer(new Worksheet(sheet));
                }
                catch (...)
                {
                    return XLSXCpp::IWorksheet<>::empty();
                }
            }

            virtual bool removeSheet(std::string _Name) override
            {
                if(!m_Document.isOpen())
                    return false;

                try
                {
                    m_Document.workbook().deleteSheet(_Name);
                    return true;
                }
                catch (...)
                {
                    return false;
                }
            }

            virtual std::vector<WorksheetPointer> Sheets() const override
            {
                if(!m_Document.isOpen())
                    return std::vector<WorksheetPointer>();

                std::vector<WorksheetPointer> sheets;

                for(int i = 0; i < m_Document.workbook().sheetCount(); i++)
                    sheets.push_back(findSheet(i));

                return sheets;
            }

            virtual int sheetsCount() const override
            {
                return m_Document.isOpen() ? m_Document.workbook().sheetCount() : 0;
            }

            virtual bool insertSheets(int _From, int _Count) override
            {
                if(!m_Document.isOpen())
                    return false;

                std::set<std::string> seed;
                std::map<int, std::string> newNames;
                auto sheets = Sheets();

                for(auto& sheet : sheets)
                    seed.insert(sheet->getName());

                for(int i = 0; i < _Count; i++)
                {
                    int postfix = sheetsCount();
                    std::string name = "List-" + std::to_string(postfix);

                    while (seed.contains(name))
                        name = "List-" + std::to_string(++postfix);

                    std::cout << name << "\n";

                    newNames.insert({_From + i, name});
                    seed.insert(name);
                }

                for(auto& name : newNames)
                {
                    auto sheet = addSheet(name.second);
                    sheet->setIndex(name.first);
                }

                return false;
            }

            virtual bool removeSheets(int _From, int _Count) override
            {
                if(!m_Document.isOpen())
                    return false;

                auto sheets = Sheets();

                for(size_t i = _From; i < std::min(_From + _Count, sheetsCount()); i++)
                    removeSheet(findSheet(i)->getName());

                return true;
            }

        protected:

            const ::OpenXLSX::XLDocument& m_Document;
        };

        // Document
        class Document : public XLSXCpp::IDocument
        {
        public:

            Document(){}

            virtual ~Document()
            {
                if(m_Document.isOpen())
                    m_Document.close();
            }

            // interface
            virtual bool open(const std::string& _Path) override
            {
                try
                {
                    m_Document.open(_Path);
                }
                catch(...)
                {
                    return false;
                }

                return true;
            }

            virtual void close() override
            {
                try
                {
                    if(m_Document.isOpen())
                        m_Document.close();
                }
                catch(...)
                {
                }
            }

            virtual bool save() override
            {
                try
                {
                    if(m_Document.isOpen())
                        m_Document.save();
                }
                catch(...)
                {
                    return false;
                }

                return true;
            }

            virtual bool saveAs(const std::string& _Path) override
            {
                try
                {
                    if(m_Document.isOpen())
                        m_Document.saveAs(_Path, true);
                }
                catch(...)
                {
                    return false;
                }

                return true;
            }

            virtual WorkbookPointer WorkBook() const override
            {
                return WorkbookPointer(new Workbook(m_Document));
            }

        protected:

            ::OpenXLSX::XLDocument m_Document;
        };
    }
};

#endif // XLSX_OPENXLSX_H
