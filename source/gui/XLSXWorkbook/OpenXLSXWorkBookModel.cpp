#include "OpenXLSXWorkBookModel.h"

OpenXLSXWorkBookModel::OpenXLSXWorkBookModel(
    QString  _Path,
    QObject* _Parent ) :
    QAbstractListModel(_Parent)
{
    m_Documnet->open(QFileInfo(_Path).filePath().toLocal8Bit().toStdString());

    m_PseudoRandomNumberGenerator.seed((uint_fast64_t)this);
}

OpenXLSXWorkBookModel::~OpenXLSXWorkBookModel()
{
    m_Documnet->close();
}

QVariant OpenXLSXWorkBookModel::data(const QModelIndex &index, int role) const
{
    (void)index;
    (void)role;

    if(role == Qt::DisplayRole)return QString::fromStdString(m_Documnet->WorkBook()->findSheet(index.row())->getName());

    if (!index.isValid() ||
        index.row() >= rowCount() ||
        (role != Qt::DisplayRole && role != Qt::EditRole))
        return QVariant();

    return QString::fromStdString(m_Documnet->WorkBook()->findSheet(index.row())->getName());
}

bool OpenXLSXWorkBookModel::setData(const QModelIndex &index, const QVariant &value, int role)
{
    if(!index.isValid()      ||
        role != Qt::EditRole ||
        index.row() >= (int)m_Documnet->WorkBook()->sheetsCount())
        return QAbstractListModel::setData(index, value, role);

    m_Documnet->WorkBook()->findSheet(index.row())->setName(value.toString().toStdString());

    return true;
}

int OpenXLSXWorkBookModel::rowCount(const QModelIndex&) const
{
    return m_Documnet->WorkBook()->sheetsCount();
}

bool OpenXLSXWorkBookModel::insertRows(int row, int count, const QModelIndex &parent)
{
    if(row < 0 || row >= rowCount() || count <= 0)
        return true;

    emit layoutAboutToBeChanged();

    // append sheets at the end
    beginInsertRows(parent, row, row + count - 1);

    if(!m_Documnet->WorkBook()->insertSheets(row, count))
        return false;

    endInsertRows();

    emit layoutChanged();

    return true;
}

bool OpenXLSXWorkBookModel::removeRows(int row, int count, const QModelIndex &parent)
{
    if(row < 0 || row >= rowCount() || count <= 0)
        return true;

    emit layoutAboutToBeChanged();

    beginRemoveRows(parent, row, row + count - 1);

    if(!m_Documnet->WorkBook()->removeSheets(row, count))
        return false;

    endRemoveRows();

    emit layoutChanged();

    return true;
}

bool OpenXLSXWorkBookModel::moveRows(const QModelIndex &sourceParent, int from, int count, const QModelIndex &destinationParent, int to)
{
    if(rowCount()    <= 0 ||
        from >= rowCount() ||
        to >=  rowCount()  ||
        from < 0           ||
        to < 0)
        return false;

    emit layoutAboutToBeChanged();

    if (from + 1 == to)
        std::swap(from, to);

    if(!beginMoveRows(sourceParent, from, from + count - 1, destinationParent, to))
        return false;

    if(!m_Documnet->WorkBook()->moveSheets(from, count, to))
        return false;

    endMoveRows();

    emit layoutChanged();

    return true;
}

Qt::ItemFlags OpenXLSXWorkBookModel::flags(const QModelIndex &index) const
{
    return QAbstractListModel::flags(index) | Qt::ItemIsEditable;
}

bool OpenXLSXWorkBookModel::open(QFileInfo _Info)
{
    return m_Documnet->open(_Info.absoluteFilePath().toLocal8Bit().toStdString());
}

bool OpenXLSXWorkBookModel::saveAs(QFileInfo _Info)
{
    return m_Documnet->saveAs(_Info.absoluteFilePath().toLocal8Bit().toStdString());
}

bool OpenXLSXWorkBookModel::save()
{
    return m_Documnet->save();
}

void OpenXLSXWorkBookModel::close()
{
    m_Documnet->close();
}

XLSX::WorkbookPointer OpenXLSXWorkBookModel::WorkBook()
{
    return m_Documnet->WorkBook();
}
