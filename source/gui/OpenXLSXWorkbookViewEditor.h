#ifndef OPENXLSXWORKBOOKVIEWEDITOR_H
#define OPENXLSXWORKBOOKVIEWEDITOR_H

#include <OpenXLSXWorkBookModel.h>
#include <OpenXLSXSheetViewEditor.h>

// Qt
#include <QWidget>
#include <QListView>
#include <QHeaderView>
#include <QToolBar>
#include <QFileInfo>
#include <QSplitter>
#include <QFileDialog>
#include <QMessageBox>

// OpenXLSXWorkbookViewEditor
class OpenXLSXWorkbookViewEditor : public QWidget
{
    Q_OBJECT

public:

    // constructors
    OpenXLSXWorkbookViewEditor(
        QString  _Path   = QString(),
        QWidget* _Parent = nullptr);

    // virtual destructor
    virtual ~OpenXLSXWorkbookViewEditor();

protected:

    QListView* m_List = nullptr;

    // actions
    QAction* m_AddRowAction       = nullptr;
    QAction* m_RemoveRowAction    = nullptr;
    QAction* m_MoveUpAction       = nullptr;
    QAction* m_MoveDownAction     = nullptr;

    QAction* m_SaveAction   = nullptr;
    QAction* m_SaveAsAction = nullptr;
    QAction* m_OpenAction   = nullptr;

    // service methods
    QWidget* CreateTable(const QModelIndex &index);
    void OpenXLSXFile(QString _Path);
    void RemoveAllSheetEditors();
};

#endif // OPENXLSXWORKBOOKVIEWEDITOR_H
