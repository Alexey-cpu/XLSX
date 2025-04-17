#include "OpenXLSXSheetModel.h"

OpenXLSXSheetModel::OpenXLSXSheetModel(XLSX::IWorksheet<>::WorksheetPointer _Sheet, QObject* _Parent) :
    QAbstractTableModel(_Parent),
    m_Sheet(_Sheet)
{
    // tweak values in cells
    for(int i = 0 ; i < m_Sheet->rowCount(); i++)
    {
        for(int j = 0 ; j < m_Sheet->columnCount(); j++)
        {
            QString stringValue =
                QString::fromStdString(m_Sheet->getData<std::string>(i, j));

            if(stringValue.isEmpty())
                continue;

            bool   valueConverted = false;
            double numbericValue  = stringValue.toDouble(&valueConverted);

            if(valueConverted)
                m_Sheet->setData<double>(i, j, numbericValue);
        }
    }
}

OpenXLSXSheetModel::~OpenXLSXSheetModel(){}

QVariant OpenXLSXSheetModel::data(const QModelIndex &index, int role) const
{
    if( !index.isValid() ||
        (role != Qt::DisplayRole && role != Qt::EditRole) ||
        index.row() >= (int)m_Sheet->rowCount() ||
        index.column() >= (int)m_Sheet->columnCount() )
        return QVariant();

    return QString::fromStdString(m_Sheet->getData<std::string>(index.row(), index.column())); // display as string !!!
}

bool OpenXLSXSheetModel::setData(const QModelIndex &index, const QVariant &value, int role)
{
    if( !index.isValid() ||
        role != Qt::EditRole ||
        index.row() >= (int)m_Sheet->rowCount() ||
        index.column() >= (int)m_Sheet->columnCount() )
        return QAbstractTableModel::setData(index, value, role);

    bool   valueConverted = false;
    double numericValue   = value.toDouble(&valueConverted);

    if(valueConverted)
        m_Sheet->setData(index.row(), index.column(), numericValue);
    else
        m_Sheet->setData(index.row(), index.column(), value.toString().toStdString());

    return true;
}

QVariant OpenXLSXSheetModel::headerData(int section, Qt::Orientation orientation, int role) const
{
    if(section < 0 || role != Qt::DisplayRole)
        return QAbstractTableModel::headerData(section, orientation, role);

    if(orientation == Qt::Orientation::Vertical && section < rowCount())
        return section + 1;

    //if(orientation == Qt::Orientation::Horizontal && section < columnCount())
    //    return QString::fromStdString(m_Sheet.cell(1, section + 1).getString());

    return QAbstractTableModel::headerData(section, orientation, role);
}

bool OpenXLSXSheetModel::setHeaderData(int section, Qt::Orientation orientation, const QVariant &value, int role)
{
    if(section < 0)
        return false;

    if(Qt::Orientation::Vertical && section >= rowCount())
        return false;

    if(Qt::Orientation::Horizontal && section >= columnCount())
        return false;

    return QAbstractTableModel::setHeaderData(section, orientation, value, role);
}

int OpenXLSXSheetModel::rowCount(const QModelIndex&) const
{    
    return m_Sheet->rowCount();
}

int OpenXLSXSheetModel::columnCount(const QModelIndex&) const
{
    return m_Sheet->columnCount();
}

bool OpenXLSXSheetModel::insertRows(int row, int count, const QModelIndex &parent)
{
    // tweak starting row
    if(row > rowCount())
        row = rowCount();

    // append rows to the end
    beginInsertRows(parent, rowCount(), rowCount() + count);

    if(!m_Sheet->insertRows(row, count))
        return false;

    endInsertRows();

    emit layoutChanged();

    return true;
}

bool OpenXLSXSheetModel::insertColumns(int column, int count, const QModelIndex &parent)
{
    // tweak starting column
    if(column > columnCount())
        column = columnCount();

    // append columns to the end
    beginInsertColumns(parent, columnCount(), columnCount() + count);

    if(!m_Sheet->insertColumns(column, count))
        return false;

    endInsertColumns();

    return true;
}

bool OpenXLSXSheetModel::moveRows(const QModelIndex &sourceParent, int from, int count, const QModelIndex &destinationParent, int to)
{
    if(to < 0 || from < 0 || to >= rowCount() || from >= rowCount())
        return false;

    if (from + 1 == to)
        std::swap(from, to);

    if(!beginMoveRows(QModelIndex(), from, from + count - 1, QModelIndex(), to))
        return false;

    // retrieve source values
    if(!m_Sheet->moveRows(from, count, to))
        return false;

    endMoveRows();

    return true;
}

bool OpenXLSXSheetModel::moveColumns(const QModelIndex &sourceParent, int from, int count, const QModelIndex &destinationParent, int to)
{
    if(to < 0 || from < 0 || to >= columnCount() || from >= columnCount())
        return false;

    if(from + 1 == to)
        std::swap(from, to);

    if(!beginMoveColumns(QModelIndex(), from, from + count - 1, QModelIndex(), to))
        return false;

    if(!m_Sheet->moveColumns(from, count, to))
        return false;

    endMoveColumns();

    return true;
}

void OpenXLSXSheetModel::sort(int column, Qt::SortOrder order)
{
    emit layoutAboutToBeChanged();

    // TODO: implement this !!!

    emit layoutChanged();
}

Qt::ItemFlags OpenXLSXSheetModel::flags(const QModelIndex &index) const
{
    return QAbstractTableModel::flags(index) | Qt::ItemIsEditable;
}
