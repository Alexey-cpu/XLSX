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

    int m_LeadingColumn;

    // service methods
    void OnAddRowAction(bool);
    void OnAddColumnAction(bool);
    void OnMoveUpAction(bool);
    void OnMoveDownAction(bool);
    void OnMoveLeftAction(bool);
    void OnMoveRightAction(bool);
};

#endif // OPENXLSXSHEETVIEWEDITOR_H
