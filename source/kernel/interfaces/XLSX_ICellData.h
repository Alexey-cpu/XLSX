#ifndef XLSX_ICELLDATA_H
#define XLSX_ICELLDATA_H

namespace XLSX
{
    template<typename __type>
    class ICellData
    {
    public:

        // constructors
        ICellData(){}

        // virtual destructor
        virtual ~ICellData(){}

        // getters
        virtual __type getData(int _Row, int _Column, ICellData<__type>* _Data = nullptr) const = 0;

        // setters
        virtual bool setData(int _Row, int _Column, __type _Value, ICellData<__type>* _Data = nullptr) = 0;
    };
};

#endif // XLSX_ICELLDATA_H
