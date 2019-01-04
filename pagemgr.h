#ifndef PAGEMGR_H
#define PAGEMGR_H

#include <QObject>

class QWebEnginePage;
class QWebEngineView;

struct  Data {
    QString date;
    QString code;
    QString num;
    QString amount;

    QString sheetname;
    int     row;
};

class PageMgr : public QObject
{
    Q_OBJECT
public:
    explicit PageMgr(QObject *parent = nullptr);
    ~PageMgr();

    void init();

public slots:
    void onLoadFinished(bool);
    void onFileLoadFinished(QString);
    void onScreenshoted(int, int, int, int);

    bool readXlsx();

private:
    void gotoNext();

    QWebEnginePage *page() const;

private:
    QWebEngineView *view;
    QList<Data> list;
    QString filepath;
    QString fileName;
};

#endif // PAGEMGR_H
