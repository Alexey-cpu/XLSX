#include "OpenXLSXWorkBookModel.h"

OpenXLSXWorkBookModel::OpenXLSXWorkBookModel(
    QString  _Path,
    QObject* _Parent ) :
    QAbstractListModel(_Parent)
{
    m_Doc->open(QFileInfo(_Path).filePath().toLocal8Bit().toStdString());

    m_PseudoRandomNumberGenerator.seed((uint_fast64_t)this);
}

OpenXLSXWorkBookModel::~OpenXLSXWorkBookModel()
{
    m_Doc->close();
}

QVariant OpenXLSXWorkBookModel::data(const QModelIndex &index, int role) const
{
    (void)index;
    (void)role;

    if (!index.isValid() ||
        index.row() >= (int)m_Doc->WorkBook()->sheetsCount() ||
        (role != Qt::DisplayRole && role != Qt::EditRole))
        return QVariant();

    return QString::fromStdString(m_Doc->WorkBook()->findSheet(index.row())->getName());
}

bool OpenXLSXWorkBookModel::setData(const QModelIndex &index, const QVariant &value, int role)
{
    if(!index.isValid()      ||
        role != Qt::EditRole ||
        index.row() >= (int)m_Doc->WorkBook()->sheetsCount())
        return QAbstractListModel::setData(index, value, role);

    m_Doc->WorkBook()->findSheet(index.row())->setName(value.toString().toStdString());

    return true;
}

int OpenXLSXWorkBookModel::rowCount(const QModelIndex&) const
{
    return  m_Doc->WorkBook()->sheetsCount();
}

int OpenXLSXWorkBookModel::columnCount(const QModelIndex&) const
{
    return 1;
}

bool OpenXLSXWorkBookModel::insertRows(int row, int count, const QModelIndex &parent)
{
    if(row < 0 || row >= rowCount() || count <= 0)
        return true;

    // append sheets at the end
    beginInsertRows(parent, rowCount(), rowCount() + count);

    if(!m_Doc->WorkBook()->insertSheets(row, count))
        return false;

    endInsertRows();

    // move added sheets to an appropariate index
    return true;
}

bool OpenXLSXWorkBookModel::removeRows(int row, int count, const QModelIndex &parent)
{
    if(row < 0 || row >= rowCount() || count <= 0)
        return true;

    beginRemoveRows(parent, row, row + count);

    if(!m_Doc->WorkBook()->removeSheets(row, count))
        return false;

    endRemoveRows();

    return true;
}

bool OpenXLSXWorkBookModel::moveRows(
    const QModelIndex &sourceParent,
    int from,
    int count,
    const QModelIndex &destinationParent,
    int to )
{
    if(rowCount()    <= 0 ||
        columnCount() <= 0 ||
        from >= rowCount() ||
        to >=  rowCount()  ||
        from < 0           ||
        to < 0)
        return false;

    if (from + 1 == to)
        std::swap(from, to);

    if(!beginMoveRows(sourceParent, from, from + count - 1, destinationParent, to))
        return false;

    // TODO: implemente this !!!

    endMoveRows();

    return true;
}

Qt::ItemFlags OpenXLSXWorkBookModel::flags(const QModelIndex &index) const
{
    return QAbstractListModel::flags(index) | Qt::ItemIsEditable;
}
