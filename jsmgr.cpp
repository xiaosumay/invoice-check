#include "jsmgr.h"

JsMgr::JsMgr(QObject *parent) : QObject(parent)
{

}

void JsMgr::trigger(QString cmd)
{
    emit loadFinished(cmd);
}

void JsMgr::screenshot(int left, int top, int width, int height)
{
    emit screenshoted(left, top, width, height);
}
