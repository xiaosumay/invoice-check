#include "pagemgr.h"

#include <QApplication>
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
    os << "(" << data.date << "," << data.code << "," << data.num << "," << data.amount << "," << data.sheetname << "," << data.row << ")";
    return os;
}

bool PageMgr::readXlsx()
{
    fileName = QFileDialog::getOpenFileName(0, tr("Open Excel"), QDir::homePath(), tr("Excel Files (*.xlsx)"));

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

            data.sheetname = name;
            data.row = i;

            auto date = xlsx.cellAt(i, 3);
            if(date) {
                if(date->isDateTime())
                    data.date = date->dateTime().toString("yyyyMMdd");
                else
                    data.date = date->value().toString();
            }

            auto code = xlsx.cellAt(i, 4);
            if(code)
                data.code = code->value().toString();

            auto num = xlsx.cellAt(i, 5);
            if(num)
                data.num = num->value().toString();

            auto amount = xlsx.cellAt(i, 6);
            if(amount)
                data.amount = amount->value().toString();

            if(data.date.isEmpty() || data.code.isEmpty() || data.num.isEmpty() || data.amount.isEmpty())
                break;

            auto check = xlsx.cellAt(i, 8);

            if(!check || check->value().toString().isEmpty()) {
                list.append(data);
            }
        }
    }

    return true;
}

void PageMgr::gotoNext()
{
    if(list.isEmpty()) {
        view->hide();
        QMessageBox::information(0, "Title", QStringLiteral("文件解析结束！自动关闭！"));
        qApp->quit();
        return;
    }

    QXlsx::Document xlsx(fileName);

    auto data = list.first();

    if(xlsx.selectSheet(data.sheetname)) {
        qDebug() << xlsx.currentWorksheet()->sheetName();
    }


    auto format = xlsx.rowFormat(data.row);
    format.setFontColor(Qt::red);
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    xlsx.write(data.row, 8, QStringLiteral("已查验"), format);

    while(!xlsx.save()) {
        QMessageBox::critical(0, "error", QStringLiteral("文件占用，关闭其他使用此 %1.xlsx 的程序!").arg(QFileInfo(fileName).baseName()));
    }

    list.removeFirst();
    view->reload();
}

QWebEnginePage *PageMgr::page() const
{
    return view->page();
}
