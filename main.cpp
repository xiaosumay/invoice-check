#include <QApplication>
#include "pagemgr.h"

int main(int argc, char *argv[])
{
    QCoreApplication::setAttribute(Qt::AA_EnableHighDpiScaling);
    QApplication a(argc, argv);

    PageMgr mgr;

    if(!mgr.readXlsx()) {
        return 0;
    }

    mgr.init();

    return a.exec();
}
