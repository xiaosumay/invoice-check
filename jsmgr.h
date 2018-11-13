#ifndef JSMGR_H
#define JSMGR_H

#include <QObject>

class JsMgr : public QObject
{
    Q_OBJECT
public:
    explicit JsMgr(QObject *parent = nullptr);

signals:
    void loadFinished(QString cmd);
    void screenshoted(int left, int top, int width, int height);

public slots:
    void trigger(QString cmd);
    void screenshot(int left, int top, int width, int height);

};

#endif // JSMGR_H
