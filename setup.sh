#!/bin/bash

# 配置脚本：快速替换 manifest.xml 中的占位符

echo "================================"
echo "Excel AI 插件 GitHub Pages 配置"
echo "================================"
echo ""

# 获取用户输入
read -p "请输入你的 GitHub 用户名: " username
read -p "请输入你的仓库名 (默认: excel-ai-cleaner): " repo
repo=${repo:-excel-ai-cleaner}

echo ""
echo "配置信息："
echo "  GitHub 用户名: $username"
echo "  仓库名: $repo"
echo "  插件地址: https://$username.github.io/$repo/"
echo ""

# 确认
read -p "确认以上信息正确？(y/n): " confirm
if [[ $confirm != "y" && $confirm != "Y" ]]; then
    echo "已取消操作"
    exit 1
fi

# 替换 manifest.xml 中的占位符
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS
    sed -i '' "s/YOUR_USERNAME/$username/g" manifest.xml
    sed -i '' "s/YOUR_REPO/$repo/g" manifest.xml
else
    # Linux
    sed -i "s/YOUR_USERNAME/$username/g" manifest.xml
    sed -i "s/YOUR_REPO/$repo/g" manifest.xml
fi

echo ""
echo "✅ manifest.xml 已更新！"
echo ""
echo "下一步操作："
echo "1. 创建 Git 仓库并推送到 GitHub："
echo "   git init"
echo "   git add ."
echo "   git commit -m 'Initial commit'"
echo "   git remote add origin https://github.com/$username/$repo.git"
echo "   git branch -M main"
echo "   git push -u origin main"
echo ""
echo "2. 在 GitHub 仓库设置中启用 Pages（Settings > Pages > main branch）"
echo ""
echo "3. 在 Excel 中加载插件"
echo "   地址: https://$username.github.io/$repo/manifest.xml"
