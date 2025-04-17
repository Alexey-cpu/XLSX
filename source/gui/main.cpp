#include <QApplication>

#include <OpenXLSXWorkbookViewEditor.h>

int main(int argc, char *argv[])
{
    // initialize application
    QApplication application(argc, argv);

// force numeric locale settings
#ifdef __APPLE__
    setlocale(LC_NUMERIC,"C");
#else
    std::setlocale(LC_NUMERIC,"C");
#endif

    QLocale::setDefault(QLocale::C);

    // create main window
    OpenXLSXWorkbookViewEditor powerCat;
    powerCat.showMaximized();
    return application.exec();
}
