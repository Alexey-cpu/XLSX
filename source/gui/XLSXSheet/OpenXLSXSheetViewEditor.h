#ifndef OPENXLSXSHEETVIEWEDITOR_H
#define OPENXLSXSHEETVIEWEDITOR_H

#include <OpenXLSXSheetModel.h>

// Qt
#include <QWidget>
#include <QTableView>
#include <QHeaderView>
#include <QToolBar>
#include <QDialog>
#include <QLineEdit>
#include <QVBoxLayout>
#include <QPushButton>
#include <QFormLayout>
#include <QCheckBox>
#include <QMouseEvent>

class TableView : public QTableView
{
    Q_OBJECT

public:

    TableView(XLSX::WorksheetPointer _Sheet, QWidget* _Parent = nullptr) :
        QTableView(_Parent)
    {
        // create table view
        this->setModel(new OpenXLSXSheetModel(_Sheet, this));
        resizeColumnsToContents();
        horizontalHeader()->setStretchLastSection(true);
        setEditTriggers(QTableView::DoubleClicked);

        connect(
            this,
            &QTableView::pressed,
            this,
            [this](const QModelIndex &index)
            {
                m_ReferenceCellData = index.data();
            }
            );

        connect(
            selectionModel(),
            &QItemSelectionModel::selectionChanged,
            this,
            [this](const QItemSelection &selected, const QItemSelection &deselected)
            {
                m_ReferenceColumn = -1;
            }
        );

        connect(
            horizontalHeader(),
            &QHeaderView::sectionClicked,
            this,
            [this](int logicalIndex)
            {
                m_ReferenceColumn = logicalIndex;
            }
            );

        connect(
            horizontalHeader(),
            &QHeaderView::sectionEntered,
            this,
            [this](int logicalIndex)
            {
                m_ReferenceColumn = logicalIndex;
            }
            );

        connect(
            horizontalHeader(),
            &QHeaderView::sectionResized,
            this,
            [this](int logicalIndex, int oldSize, int newSize)
            {
                if(m_ReferenceColumn < 0)
                    m_ReferenceColumn = logicalIndex;

                auto selection = selectionModel()->selectedColumns();

                for(auto& sel : selection)
                {
                    horizontalHeader()->resizeSection(
                        sel.column(),
                        horizontalHeader()->sectionSize(m_ReferenceColumn)
                        );
                }
            }
            );
    }

    virtual ~TableView(){}

    virtual void mouseReleaseEvent(QMouseEvent *event) override
    {
        OpenXLSXSheetModel* model = findChild<OpenXLSXSheetModel*>();

        if(model == nullptr)
            return;


        if(event->button() == Qt::MouseButton::LeftButton &&
            (bool)(event->modifiers() & Qt::KeyboardModifier::ControlModifier))
        {
            model->emitLayoutAboutToBeChanged();

            auto selection = selectionModel()->selectedIndexes();

            for(auto& index : selection)
                model->setData(index, m_ReferenceCellData);

            model->emitLayoutChanged();
        }

        QTableView::mouseReleaseEvent(event);
    }

protected:

    int m_ReferenceColumn = -1;

    QVariant m_ReferenceCellData;
};

// OpenXLSXSheetViewEditor
class OpenXLSXSheetViewEditor : public QWidget
{
    Q_OBJECT

public:

    // constructors
    OpenXLSXSheetViewEditor(
        XLSX::WorksheetPointer _Sheet,
        QWidget* _Parent = nullptr );

    // virtual destructor
    virtual ~OpenXLSXSheetViewEditor();

protected:

    // views
    QTableView* m_TableView = nullptr;
    QToolBar*   m_ToolBar   = nullptr;

    // actions
    QAction* m_AddRowAction     = nullptr;
    QAction* m_AddColumnAction  = nullptr;
    QAction* m_MoveUpAction     = nullptr;
    QAction* m_MoveDownAction   = nullptr;
    QAction* m_MoveLeftAction   = nullptr;
    QAction* m_MoveRightAction  = nullptr;

    // service methods
    void OnAddRowAction(bool);
    void OnAddColumnAction(bool);
    void OnMoveUpAction(bool);
    void OnMoveDownAction(bool);
    void OnMoveLeftAction(bool);
    void OnMoveRightAction(bool);
};

#endif // OPENXLSXSHEETVIEWEDITOR_H
