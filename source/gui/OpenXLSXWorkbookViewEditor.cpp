#include "OpenXLSXWorkbookViewEditor.h"

// OpenXLSXWorkbookViewEditor
OpenXLSXWorkbookViewEditor::OpenXLSXWorkbookViewEditor(
    QString  _Path,
    QWidget* _Parent) :
    QWidget(_Parent)
{
    // create list view
    m_List = new QListView(this);

    connect(
        m_List,
        &QListView::pressed,
        this,
        [this](const QModelIndex &index)
        {
            Q_UNUSED(index)

            QSplitter* splitter =
                this->findChild<QSplitter*>();

            if(splitter == nullptr)
                return;

            if(splitter->count() <= 1)
            {
                QWidget* widget = CreateTable(index);

                if( widget != nullptr )
                {
                    splitter->addWidget( widget );
                    widget->show();
                }
            }
            else
            {
                QWidget* widget = CreateTable(index);

                if(widget != nullptr)
                {
                    QWidget* replaced = splitter->replaceWidget(1, widget);
                    widget->show();

                    if(replaced != nullptr)
                        delete replaced;
                }
            }

            splitter->setSizes(
                {
                    splitter->rect().size().width() * 1 / 3,
                    splitter->rect().size().width() * 2 / 2
                }
                );
        }
    );

    // create actions
    m_AddRowAction    = new QAction("AddRow", this );
    m_RemoveRowAction = new QAction("RemoveRow", this );
    m_MoveUpAction    = new QAction("MoveUp", this );
    m_MoveDownAction  = new QAction("MoveDown", this );
    m_SaveAction      = new QAction("Save", this );
    m_SaveAsAction    = new QAction("SaveAs", this );
    m_OpenAction      = new QAction("Open", this );

    // m_MoveLeftAction
    connect(
        m_MoveUpAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            if( m_List == nullptr ||
                m_List->selectionModel() == nullptr )
                return;

            OpenXLSXWorkBookModel* model = this->findChild<OpenXLSXWorkBookModel*>();
            auto selection = m_List->selectionModel()->selectedRows();

            if(model != nullptr && !selection.empty())
                model->moveRow( QModelIndex(), selection.first().row(), QModelIndex(), selection.first().row() - 1 );
        }
        );

    // m_MoveDownAction
    connect(
        m_MoveDownAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            if( m_List == nullptr ||
                m_List->selectionModel() == nullptr )
                return;

            OpenXLSXWorkBookModel* model = this->findChild<OpenXLSXWorkBookModel*>();
            auto selection = m_List->selectionModel()->selectedRows();

            if(model != nullptr && !selection.empty())
                model->moveRow(QModelIndex(), selection.first().row(), QModelIndex(), selection.first().row() + 1);
        }
        );

    // m_AddRowAction
    connect(
        m_AddRowAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            OpenXLSXWorkBookModel* model =
                this->findChild<OpenXLSXWorkBookModel*>();

            if(m_List == nullptr ||
                model  == nullptr ||
                m_List->selectionModel() == nullptr )
                return;

            auto selection = m_List->selectionModel()->selectedRows();

            model->insertRows(
                selection.empty() ? std::max(model->rowCount() - 1, 0) : selection.first().row(),
                1
            );
        }
        );

    // m_RemoveRowAction
    connect(
        m_RemoveRowAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            if( m_List == nullptr ||
                m_List->selectionModel() == nullptr )
                return;

            OpenXLSXWorkBookModel* model = this->findChild<OpenXLSXWorkBookModel*>();
            auto selection = m_List->selectionModel()->selectedRows();

            if(model == nullptr || selection.empty())
                return;

            QMessageBox::StandardButton answer =
                QMessageBox::question(
                    this,
                    "Warning",
                    QString("Are you sure you would like to remove sheet {%1} ?")
                        .arg(QString::fromStdString(model->Doc()->WorkBook()->findSheet(selection.first().row() + 1)->getName()))
                    );

            if(answer == QMessageBox::StandardButton::Yes)
            {
                RemoveAllSheetEditors();
                model->removeRows(selection.first().row(), 1);
            }
        }
        );

    // m_SaveAction
    connect(
        m_SaveAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            OpenXLSXWorkBookModel* model =
                this->findChild<OpenXLSXWorkBookModel*>();

            if(model != nullptr)
                model->Doc()->save();
        }
        );

    // m_SaveAsAction
    connect(
        m_SaveAsAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            OpenXLSXWorkBookModel* model =
                this->findChild<OpenXLSXWorkBookModel*>();

            if(model == nullptr)
                return;

            QString path = QFileDialog::getSaveFileName(this);

            model->Doc()->saveAs(path.toStdString());
        }
        );

    // m_OpenAction
    connect(
        m_OpenAction,
        &QAction::triggered,
        m_List,
        [this]()
        {
            QString path = QFileDialog::getOpenFileName(this);

            if( QFileInfo(path).suffix() != "xlsx" &&
                QFileInfo(path).suffix() != "XLSX" )
            {
                QMessageBox::critical(
                    nullptr,
                    "Error",
                    QString("Trying to open file of unsupported format %1").arg( QFileInfo(path).suffix()));

                return;
            }

            OpenXLSXFile(path);
        }
        );

    // add actions to tool bar
    QToolBar* toolBar = new QToolBar(this);
    toolBar->addAction(m_AddRowAction);
    toolBar->addAction(m_RemoveRowAction);
    toolBar->addSeparator();
    toolBar->addAction(m_MoveUpAction);
    toolBar->addAction(m_MoveDownAction);
    toolBar->addSeparator();
    toolBar->addAction(m_SaveAction);
    toolBar->addAction(m_SaveAsAction);
    toolBar->addAction(m_OpenAction);

    // create object names for some QActions widgets
    toolBar->widgetForAction(m_AddRowAction)->setObjectName("ActionAddRow");
    toolBar->widgetForAction(m_RemoveRowAction)->setObjectName("ActionRemoveRow");
    toolBar->widgetForAction(m_MoveUpAction)->setObjectName("ActionMoveUp");
    toolBar->widgetForAction(m_MoveDownAction)->setObjectName("ActionMoveDown");
    toolBar->widgetForAction(m_SaveAction)->setObjectName("ActionSave");
    toolBar->widgetForAction(m_SaveAsAction)->setObjectName("ActionSaveAs");
    toolBar->widgetForAction(m_OpenAction)->setObjectName("ActionOpen");

    // create splitter
    QSplitter* splitter = new QSplitter( this );
    splitter->addWidget(m_List);

    // layouting
    QVBoxLayout* layout = new QVBoxLayout(this);
    layout->addWidget( toolBar );
    layout->addWidget(splitter);

    // open file
    OpenXLSXFile( _Path );
}

OpenXLSXWorkbookViewEditor::~OpenXLSXWorkbookViewEditor(){}

QWidget* OpenXLSXWorkbookViewEditor::CreateTable(const QModelIndex &index)
{
    OpenXLSXWorkBookModel* model =
        this->findChild<OpenXLSXWorkBookModel*>();

    if(model == nullptr)
        return nullptr;

    std::string name = m_List->model()->data(index).toString().toStdString();

    OpenXLSXSheetViewEditor* editor =
        new OpenXLSXSheetViewEditor(model->Doc()->WorkBook()->findSheet(name));

    return editor;
}

void OpenXLSXWorkbookViewEditor::OpenXLSXFile(QString _Path)
{
    if( !QFileInfo::exists(_Path) )
        return;

    if( m_List == nullptr )
        return;

    RemoveAllSheetEditors();

    OpenXLSXWorkBookModel* model =
        m_List->findChild<OpenXLSXWorkBookModel*>();

    if( model != nullptr )
        delete model;

    m_List->setModel( new OpenXLSXWorkBookModel( _Path, m_List ) );
}

void OpenXLSXWorkbookViewEditor::RemoveAllSheetEditors()
{
    QSplitter* splitter =
        this->findChild<QSplitter*>();

    if( splitter == nullptr )
        return;

    auto editors = splitter->findChildren<OpenXLSXSheetViewEditor*>();

    for(auto& editor : editors)
    {
        if(editor != nullptr)
            delete editor;
    }
}
