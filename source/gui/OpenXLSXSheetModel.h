#ifndef OPENXLSXSHEETMODEL_H
#define OPENXLSXSHEETMODEL_H

#include <XLSX.h>

// Qt
#include <QAbstractTableModel>

// OpenXLSXSheetModel
class OpenXLSXSheetModel : public QAbstractTableModel
{
    Q_OBJECT

public:

    // constructors
    OpenXLSXSheetModel(XLSX::IWorksheet<>::WorksheetPointer _Sheet, QObject* _Parent = nullptr);

    // vrtual destructor
    virtual ~OpenXLSXSheetModel();

    // QAbstractTableModel
    virtual QVariant data(const QModelIndex& _Index, int _Role = Qt::DisplayRole) const override;
    virtual bool setData(const QModelIndex& _Index, const QVariant& _Value, int _Role = Qt::EditRole) override;
    virtual QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
    virtual bool setHeaderData(int section, Qt::Orientation orientation, const QVariant& value, int role = Qt::EditRole) override;
    virtual int rowCount(const QModelIndex& = QModelIndex()) const override;
    virtual int columnCount(const QModelIndex& = QModelIndex()) const override;
    virtual bool insertRows(int row, int count, const QModelIndex& parent = QModelIndex()) override;
    virtual bool insertColumns(int column, int count, const QModelIndex& parent = QModelIndex()) override;
    virtual bool moveRows(const QModelIndex& sourceParent, int from, int count, const QModelIndex& destinationParent, int to) override;
    virtual bool moveColumns(const QModelIndex& sourceParent, int from, int count, const QModelIndex& destinationParent, int to) override;
    virtual void sort(int column, Qt::SortOrder order = Qt::AscendingOrder) override;
    virtual Qt::ItemFlags flags(const QModelIndex& index) const override;

protected:

    XLSX::IWorksheet<>::WorksheetPointer m_Sheet;
};

#endif // OPENXLSXSHEETMODEL_H
