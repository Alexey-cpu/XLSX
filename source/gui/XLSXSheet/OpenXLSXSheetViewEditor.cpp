#include "OpenXLSXSheetViewEditor.h"

#include "OpenXLSXSheetModel.h"

// OpenXLSXSheetViewEditor
OpenXLSXSheetViewEditor::OpenXLSXSheetViewEditor(
    XLSX::WorksheetPointer _Sheet,
    QWidget* _Parent) :
    QWidget(_Parent)
{
    // create table view
    m_TableView = new TableView(_Sheet, this);

    // create actions
    m_AddRowAction     = new QAction("AddRow", this );
    m_AddColumnAction  = new QAction("AddColumn", this );
    m_MoveUpAction     = new QAction("MoveUp", this );
    m_MoveDownAction   = new QAction("MoveDown", this );
    m_MoveLeftAction   = new QAction("MoveLeft", this );
    m_MoveRightAction  = new QAction("MoveRight", this );

    // m_AddRowAction
    connect(
        m_AddRowAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnAddRowAction
        );

    // m_AddColumnAction
    connect(
        m_AddColumnAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnAddColumnAction
        );

    // m_MoveUpAction
    connect(
        m_MoveUpAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnMoveUpAction
        );

    // m_MoveDownAction
    connect(
        m_MoveDownAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnMoveDownAction
        );

    // m_MoveLeftAction
    connect(
        m_MoveLeftAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnMoveLeftAction
        );

    // m_MoveRightAction
    connect(
        m_MoveRightAction,
        &QAction::triggered,
        this,
        &OpenXLSXSheetViewEditor::OnMoveRightAction
        );

    // add actions to tool bar
    m_ToolBar = new QToolBar(this);
    m_ToolBar->addAction(m_AddRowAction);
    m_ToolBar->addAction(m_AddColumnAction);
    m_ToolBar->addSeparator();
    m_ToolBar->addAction(m_MoveUpAction);
    m_ToolBar->addAction(m_MoveDownAction);
    m_ToolBar->addAction(m_MoveLeftAction);
    m_ToolBar->addAction(m_MoveRightAction);

    // create object names for some QActions widgets
    m_ToolBar->widgetForAction(m_AddRowAction)->setObjectName("ActionAddRow");
    m_ToolBar->widgetForAction(m_AddColumnAction)->setObjectName("ActionAddColumn");
    m_ToolBar->widgetForAction(m_MoveUpAction)->setObjectName("ActionMoveUp");
    m_ToolBar->widgetForAction(m_MoveDownAction)->setObjectName("ActionMoveDown");
    m_ToolBar->widgetForAction(m_MoveLeftAction)->setObjectName("ActionMoveLeft");
    m_ToolBar->widgetForAction(m_MoveRightAction)->setObjectName("ActionMoveRight");

    QVBoxLayout* layout = new QVBoxLayout(this);
    layout->addWidget(m_ToolBar);
    layout->addWidget(m_TableView);
}

OpenXLSXSheetViewEditor::~OpenXLSXSheetViewEditor(){}

void OpenXLSXSheetViewEditor::OnAddRowAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection = m_TableView->selectionModel()->selectedIndexes();

    model->insertRows((selection.empty() ? std::max(model->rowCount() - 1, 0) : selection.first().row()), 1);
}

void OpenXLSXSheetViewEditor::OnAddColumnAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection =
        m_TableView->selectionModel()->selectedIndexes();

    if(model->insertColumns(selection.empty() ? std::max(model->columnCount() - 1, 0) : selection.first().column(), 1))
    {
        // TODO: when filters pane is added here, we need to update filters
    }
}

void OpenXLSXSheetViewEditor::OnMoveUpAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection = m_TableView->selectionModel()->selectedIndexes();

    model->moveRow(
        QModelIndex(),
        selection.first().row(),
        QModelIndex(),
        selection.first().row() - 1);
}

void OpenXLSXSheetViewEditor::OnMoveDownAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection = m_TableView->selectionModel()->selectedIndexes();

    model->moveRow(
        QModelIndex(),
        selection.first().row(),
        QModelIndex(),
        selection.first().row() + 1);
}

void OpenXLSXSheetViewEditor::OnMoveLeftAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection = m_TableView->selectionModel()->selectedIndexes();

    if(selection.empty())
        return;

    model->moveColumn(QModelIndex(), selection.first().column(), QModelIndex(), selection.first().column() - 1);
}

void OpenXLSXSheetViewEditor::OnMoveRightAction(bool)
{
    OpenXLSXSheetModel* model =
        m_TableView != nullptr ?
            m_TableView->findChild<OpenXLSXSheetModel*>() :
                nullptr;

    if(model == nullptr)
        return;

    auto selection = m_TableView->selectionModel()->selectedIndexes();

    if(selection.empty())
        return;

    model->moveColumn(QModelIndex(), selection.first().column(), QModelIndex(), selection.first().column() + 1);
}
