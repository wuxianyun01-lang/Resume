# 吴志勇 · IT运维工程师 简历项目

定位：IT运维工程师（桌面运维 / 轻量开发 / AI 应用）。

## 交付物

- 网页版简历：`index.html`，可直接打开，也可部署到 GitHub Pages / Netlify。
- Word 简历：
  - `word/吴志勇-IT运维工程师-一页版.docx`
  - `word/吴志勇-IT运维工程师-两页版.docx`
- PDF 简历：与 Word 同名，位于 `word/` 目录。
- BOSS 可粘贴纯文本版：`docs/简历纯文本版.txt`
- HR 话术与关键词说明：`docs/HR话术与使用说明.md`

## 本地预览

直接双击 `index.html` 即可。也可以起一个本地服务：

```powershell
python -m http.server 8000
```

然后访问 `http://localhost:8000`。

## 修改内容

网页内容在 `index.html` 中修改；Word 内容在 `scripts/generate_resume_docx.py` 中修改，改完执行：

```powershell
python scripts/generate_resume_docx.py
```

生成后重新把 `.docx` 另存/导出为 PDF 即可。

## 部署

### GitHub Pages

1. 在 GitHub 新建仓库，并把本项目推送到 `main` 分支。
2. 仓库 Settings -> Pages -> Source 选择 `GitHub Actions`。
3. 推送后 `.github/workflows/pages.yml` 会自动构建并发布。

### Netlify

1. 打开 [Netlify Drop](https://app.netlify.com/drop)，把本项目文件夹拖进去。
2. 发布成功后把生成的链接记下来，作为备用链接。

## 照片

网页版照片位于 `assets/photo.jpg`，直接替换同名文件即可。

Word 版照片位置：右上角“照片 3.0 × 4.0 cm”占位框，用 Word/WPS 打开后把照片插入该单元格，删除占位文字即可。

## 关键词区

Word 底部有一行正常可读的【岗位关键词】，具体位置和“是否改为浅灰/透明”的说明见 `docs/HR话术与使用说明.md`。
