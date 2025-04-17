#ifndef XLSX_XLNT_WORKSHEET_H
#define XLSX_XLNT_WORKSHEET_H

// custom
#include "XLSX_IWorksheet.h"

// OpenXLSX
#include <xlnt/xlnt.hpp>

namespace XLSX
{
    namespace XLNT
    {
        class Worksheet :
            public XLSX::IWorksheet<xlnt::worksheet>,
            public XLSX::ICell<std::string>,
            public XLSX::ICell<double>,
            public XLSX::ICell<float>,
            public XLSX::ICell<int>
        {
        public:

            // constructors
            Worksheet(const xlnt::worksheet& _Sheet);

            // virtual destructor
            virtual ~Worksheet();

            // getters
            virtual std::string getName() const override;
            virtual int getIndex() const override;
            virtual std::string getData(int _Row, int _Column, ICell<std::string>* _Data = nullptr) const override;
            virtual double getData(int _Row, int _Column, ICell<double>* _Data = nullptr) const override;
            virtual float getData(int _Row, int _Column, ICell<float>* _Data = nullptr) const override;
            virtual int getData(int _Row, int _Column, ICell<int>* _Data = nullptr) const override;

            // setters
            virtual bool setName(std::string _Name) override;
            virtual bool setIndex(int _Index) override;
            virtual bool setData(int _Row, int _Column, std::string _Value, ICell<std::string>* _Data = nullptr) override;
            virtual bool setData(int _Row, int _Column, double _Value, ICell<double>* _Data = nullptr) override;
            virtual bool setData(int _Row, int _Column, float _Value, ICell<float>* _Data = nullptr) override;
            virtual bool setData(int _Row, int _Column, int _Value, ICell<int>* _Data = nullptr) override;

            // interface
            virtual int rowCount() override;
            virtual int columnCount() override;
            virtual bool insertRows(int _From, int _Count) override;
            virtual bool insertColumns(int _From, int _Count) override;
            virtual bool removeRows(int _From, int _Count) override;
            virtual bool removeColumns(int _From, int _Count) override;
            virtual bool moveRows(int _From, int _Count, int _To) override;
            virtual bool moveColumns(int _From, int _Count, int _To) override;
        };
    }
};

#endif // XLSX_XLNT_WORKSHEET_H
