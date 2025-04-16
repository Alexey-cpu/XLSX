#include "XLSX_OpenXLSX.h"

int main(int argc, char *argv[])
{
    (void)argc;
    (void)argv;

    OpenXLSX::XLDocument doc;

    XLSXCpp::OpenXLSX::Document document;

    document.open("C:/SDK/Qt_Projects/XLSX/tests/TestDocument.xlsx");

    auto workbook = document.WorkBook();

    auto sheet = workbook->findSheet(0);

    auto sheet_1 = workbook->addSheet("Sheet-1");
    sheet_1->setData<double>(0, 0, 1000);
    sheet_1->setData<float>(0, 1, 1000);
    sheet_1->setData<int>(0, 2, 1000);
    sheet_1->setData<std::string>(1,0, "Name");

    auto sheet_2 = workbook->addSheet("Sheet-2");
    sheet_2->setData<double>(0, 0, 1000);
    sheet_2->setData<float>(0, 1, 1000);
    sheet_2->setData<int>(0, 2, 1000);
    sheet_2->setData<std::string>(1,0, "Name");

    auto sheet_3 = workbook->addSheet("Sheet-3");
    sheet_3->setData<double>(0, 0, 1000);
    sheet_3->setData<float>(0, 1, 1000);
    sheet_3->setData<int>(0, 2, 1000);
    sheet_3->setData<std::string>(1,0, "Name");

    std::cout << "sheets \n";

    //if(!sheet_1->setName("Sheet-1111")) std::cout << "setName FAILED !!!" << "\n";

    auto sheets = workbook->Sheets();

    //document.close();
    //document.open("C:/SDK/Qt_Projects/XLSX/tests/TestDocument.xlsx");

    for(auto& sheet : sheets)
        std::cout << sheet->getName() << "\t" << sheet->getIndex() << "\n";

    //std::cout << "sheet->rowCount() " << sheet->rowCount() << std::endl;

    if(!sheet->removeColumns(0, 2)) std::cout << "Could not remove columns !!! \n";

    workbook->insertSheets(1, 3);
    workbook->removeSheets(2, 3);

    //std::cout << "sheet->rowCount() " << sheet->rowCount() << std::endl;

    document.saveAs("C:/SDK/Qt_Projects/XLSX/logs/TestDocument.xlsx");

    return 0;
}
