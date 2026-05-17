#!/bin/bash

cd "$(dirname "$0")"

echo "========================================"
echo "       访客登记系统"
echo "========================================"
echo ""
echo "  [1] 启动服务"
echo "  [2] 停止服务"
echo "  [3] 重启服务"
echo "  [4] 打开登记页"
echo "  [5] 打开记录页"
echo "  [6] 打开人脸识别"
echo "  [7] 打开全部页面"
echo "  [0] 退出"
echo ""

read -p "请输入选项: " num

case $num in
  1)
    if lsof -i :2312 >/dev/null 2>&1; then
      echo "服务已在运行中"
    else
      node server.js &
      sleep 1
      echo "服务已启动 http://localhost:2312"
    fi
    ;;
  2)
    pkill -f "node server.js" 2>/dev/null
    echo "服务已停止"
    ;;
  3)
    pkill -f "node server.js" 2>/dev/null
    sleep 1
    node server.js &
    sleep 1
    echo "服务已重启"
    ;;
  4)
    open http://localhost:2312/
    ;;
  5)
    open http://localhost:2312/records.html
    ;;
  6)
    open http://localhost:2312/face-scan
    ;;
  7)
    open http://localhost:2312/
    open http://localhost:2312/records.html
    open http://localhost:2312/face-scan
    ;;
  0)
    exit 0
    ;;
  *)
    echo "无效选项"
    ;;
esac