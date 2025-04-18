#ifndef OPENXLSXWORKBOOKMODEL_H
#define OPENXLSXWORKBOOKMODEL_H

#include <XLSX.h>

// Qt
#include <QAbstractListModel>
#include <QFileInfo>
#include <QMessageBox>

// OpenXLSXWorkBookModel
class OpenXLSXWorkBookModel : public QAbstractListModel
{
    Q_OBJECT

public:

    explicit OpenXLSXWorkBookModel(QString _Path, QObject* _Parent = nullptr );
    virtual ~OpenXLSXWorkBookModel();

    virtual QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override;
    virtual bool setData(const QModelIndex &index, const QVariant &value, int role = Qt::EditRole) override;
    virtual int rowCount(const QModelIndex& = QModelIndex()) const override;

    virtual bool insertRows(int row, int count, const QModelIndex &parent = QModelIndex()) override;
    virtual bool removeRows(int row, int count, const QModelIndex &parent = QModelIndex()) override;
    virtual bool moveRows(const QModelIndex &sourceParent, int from, int count, const QModelIndex &destinationParent, int to) override;
    virtual Qt::ItemFlags flags(const QModelIndex &index) const override;

    bool open(QFileInfo);
    bool saveAs(QFileInfo);
    bool save();
    QString listName(size_t);
    XLSX::WorksheetPointer findSheet(QString);

protected:


    std::shared_ptr<XLSX::IDocument> m_Documnet = std::shared_ptr<XLSX::IDocument>(new XLSX::XLNT::Document());
    std::mt19937_64 m_PseudoRandomNumberGenerator;
};

#endif // OPENXLSXWORKBOOKMODEL_H
