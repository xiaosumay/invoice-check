#include "pagemgr.h"

#include <QWebEngineView>
#include <QWebEnginePage>
#include <QWebEngineSettings>
#include <QDebug>
#include <QFileDialog>
#include <QDir>
#include <QMessageBox>
#include <QWebChannel>

#include "xlsxdocument.h"
#include "jsmgr.h"

PageMgr::PageMgr(QObject *parent)
    : QObject(parent)
    , view(new QWebEngineView)
{}

PageMgr::~PageMgr()
{
    if(view) delete view;
}

void PageMgr::init()
{
    QWebChannel *channel = new QWebChannel;
    JsMgr *jsmgr = new JsMgr;
    connect(jsmgr, &JsMgr::loadFinished, this, &PageMgr::onFileLoadFinished);
    connect(jsmgr, &JsMgr::screenshoted, this, &PageMgr::onScreenshoted);

    channel->registerObject("jsmgr", jsmgr);

    view->page()->setWebChannel(channel);

    auto setting = QWebEngineSettings::defaultSettings();

    setting->setAttribute(QWebEngineSettings::ShowScrollBars, false);
    setting->setAttribute(QWebEngineSettings::AllowRunningInsecureContent, true);

    view->setUrl(QUrl::fromUserInput("https://inv-veri.chinatax.gov.cn/"));
    view->setWindowTitle(QStringLiteral("Design for Jing.Xu"));
    view->setContextMenuPolicy(Qt::NoContextMenu);

    connect(view, &QWebEngineView::loadFinished, this, &PageMgr::onLoadFinished);

    view->showMaximized();
}

void PageMgr::onLoadFinished(bool)
{
    QFile webchannel( ":/qwebchannel.js");
    webchannel.open(QFile::ReadOnly);

    page()->runJavaScript(webchannel.readAll());

    QFile checkfp( ":/js/checkfp.js");
    checkfp.open(QFile::ReadOnly);

    page()->runJavaScript(checkfp.readAll());

    if(list.isEmpty()) {
        gotoNext();
        return;
    }

    auto data = list.first();

    auto str = QString(R"***(
        $("#fpdm").val("%1").blur()
        $("#fphm").val("%2").keyup()
        $("#kprq").val("%3").blur()
        $("#kjje").val("%4").blur()

        $("#yzm_img").click();
    )***");

    page()->runJavaScript(str.arg(data.code).arg(data.num).arg(data.date).arg(data.amount));

    qDebug() << list;
}

void PageMgr::onFileLoadFinished(QString cmd)
{
    if(cmd == QStringLiteral("reset")) {
        gotoNext();
    }
}

void PageMgr::onScreenshoted(int left, int top, int width, int height)
{
    QImage image(width, height, QImage::Format_RGB32);
    QPainter painter(&image);

    view->render(&painter, QPoint(), QRegion(left, top, width, height));

    auto data = list.first();

    image.save(QStringLiteral("%1/%2-%3-%4.png").arg(filepath).arg(data.date).arg(data.code).arg(data.num));

    gotoNext();
}

QDebug operator<<(QDebug os, Data data)
{
    os << "(" << data.date << "," << data.code << "," << data.num << "," << data.amount << ")";
    return os;
}

bool PageMgr::readXlsx()
{
    auto fileName = QFileDialog::getOpenFileName(0, tr("Open Excel"), QDir::homePath(), tr("Excel Files (*.xlsx)"));

    if(!QFileInfo(fileName).isFile() || !fileName.endsWith("xlsx", Qt::CaseInsensitive)) {
        QMessageBox::critical(0, "Error", QStringLiteral("选择xlsx文件"));
        return false;
    }

    filepath = QFileInfo(fileName).absoluteDir().canonicalPath();

    QXlsx::Document xlsx(fileName);

    list.clear();

    for(const auto &name : xlsx.sheetNames()) {
        xlsx.selectSheet(name);

        for( qint64 i = 2; i < 65536; ++i) {
            Data data{};

            auto cell = xlsx.cellAt(i, 3);
            if(cell) {
                if(cell->isDateTime())
                    data.date = cell->dateTime().toString("yyyyMMdd");
                else
                    data.date = cell->value().toString();
            }

            cell = xlsx.cellAt(i, 4);
            if(cell)
                data.code = cell->value().toString();

            cell = xlsx.cellAt(i, 5);
            if(cell)
                data.num = cell->value().toString();

            cell = xlsx.cellAt(i, 6);
            if(cell)
                data.amount = cell->value().toString();

            if(data.date.isEmpty() || data.code.isEmpty() || data.num.isEmpty() || data.amount.isEmpty())
                break;

            list.append(data);
        }
    }

    return true;
}

void PageMgr::gotoNext()
{
    if(list.isEmpty()) {
        QMessageBox::information(0, "Title", QStringLiteral("file export finished, Please close this window!"));
        return;
    }

    list.removeFirst();
    view->reload();
}

QWebEnginePage *PageMgr::page() const
{
    return view->page();
}
