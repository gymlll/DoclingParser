from pathlib import Path
import shutil
import re
import os
import pythoncom
import win32com.client as win32
from docling.document_converter import DocumentConverter
from docling_core.types.doc import ImageRefMode

def _rewrite_image_paths(md_text: str, base_dir: Path, images_dir: Path) -> str:
    pattern = re.compile(r'(!\[[^\]]*\]\()([^\)]+)(\))')
    def repl(match):
        orig_path = match.group(2)
        orig_path_resolved = (base_dir / orig_path).resolve()
        if orig_path_resolved.exists():
            target = images_dir / orig_path_resolved.name
            if not target.exists():
                shutil.move(str(orig_path_resolved), str(target))
            return f"{match.group(1)}{images_dir.name}/{target.name}{match.group(3)}"
        return match.group(0)
    return pattern.sub(repl, md_text)

def export_excel_charts(xlsx_path: Path, output_img_dir: Path) -> list[str]:
    imgs = []
    if not xlsx_path.exists():
        return imgs
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(Filename=str(xlsx_path.resolve()), ReadOnly=True, IgnoreReadOnlyRecommended=True)
        for ws in wb.Worksheets:
            ws_name = ws.Name
            exported = set()
            chart_objects = ws.ChartObjects()
            for i, chart_obj in enumerate(chart_objects, 1):
                try:
                    name = getattr(chart_obj, "Name", f"ChartObject_{i}")
                    if name in exported:
                        continue
                    ws.Activate()
                    chart_obj.Activate()
                    chart = chart_obj.Chart
                    img_name = f"{ws_name}_Chart_{i}.png"
                    img_path = str(output_img_dir / img_name)
                    chart.Export(Filename=img_path, FilterName="PNG")
                    if os.path.exists(img_path) and os.path.getsize(img_path) > 100:
                        imgs.append(img_name)
                        exported.add(name)
                except Exception:
                    pass
            try:
                shapes = ws.Shapes
                for j in range(1, shapes.Count + 1):
                    shape = shapes.Item(j)
                    if hasattr(shape, "HasChart") and shape.HasChart:
                        sname = getattr(shape, "Name", f"Shape_{j}")
                        if sname in exported:
                            continue
                        ws.Activate()
                        chart = shape.Chart
                        img_name = f"{ws_name}_ShapeChart_{j}.png"
                        img_path = str(output_img_dir / img_name)
                        chart.Export(Filename=img_path, FilterName="PNG")
                        if os.path.exists(img_path) and os.path.getsize(img_path) > 100:
                            imgs.append(img_name)
                            exported.add(sname)
            except Exception:
                pass
        wb.Close(SaveChanges=False)
        excel.Quit()
    finally:
        pythoncom.CoUninitialize()
    return imgs

def main() -> None:
    project_root = Path(__file__).resolve().parent
    template_dir = project_root / "template"
    output_root = project_root / "output"
    output_root.mkdir(parents=True, exist_ok=True)
    converter = DocumentConverter()
    xlsx_files = sorted(f for f in template_dir.rglob("*.xlsx") if not f.name.startswith("~$"))
    if not xlsx_files:
        raise FileNotFoundError(f"未找到Excel文件：{template_dir}")
    for source in xlsx_files:
        out_dir = output_root / source.stem
        images_dir = out_dir / "images"
        out_dir.mkdir(parents=True, exist_ok=True)
        images_dir.mkdir(parents=True, exist_ok=True)
        result = converter.convert(str(source))
        md_path = out_dir / "index.md"
        result.document.save_as_markdown(md_path, image_mode=ImageRefMode.REFERENCED)
        md_text = md_path.read_text(encoding="utf-8")
        md_text_rewritten = _rewrite_image_paths(md_text, md_path.parent, images_dir)
        chart_imgs = export_excel_charts(source, images_dir)
        chart_md = "\n## 嵌入式图表\n"
        if chart_imgs:
            for img in chart_imgs:
                chart_md += f"![{img}](images/{img})\n\n"
        else:
            chart_md += "- 未检测到可导出的嵌入式图表\n"
        final_md = md_text_rewritten + chart_md
        md_path.write_text(final_md, encoding="utf-8")
        shutil.copy2(source, out_dir / source.name)
        print(f"完成: {source.name} -> {out_dir} | 图表: {len(chart_imgs)}")

if __name__ == "__main__":
    main()
