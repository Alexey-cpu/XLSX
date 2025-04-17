#ifndef XLSX_ICELL_H
#define XLSX_ICELL_H

namespace XLSX
{
    template<typename __type>
    class ICell
    {
    public:

        // constructors
        ICell(){}

        // virtual destructor
        virtual ~ICell(){}

        // getters
        virtual __type getData(int _Row, int _Column, ICell<__type>* _Data = nullptr) const = 0;

        // setters
        virtual bool setData(int _Row, int _Column, __type _Value, ICell<__type>* _Data = nullptr) = 0;
    };
};

#endif // XLSX_ICELL_H
