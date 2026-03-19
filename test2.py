from pathlib import Path
import shutil
import re

from docling.document_converter import DocumentConverter
from docling_core.types.doc import ImageRefMode


def _rewrite_image_paths(md_text: str, base_dir: Path, images_dir: Path) -> str:
    pattern = re.compile(r'(!\\[[^\\]]*\\]\\()([^\\)]+)(\\))')
    def repl(match):
        orig_path = match.group(2)
        # Resolve original path relative to markdown file directory
        orig_path_resolved = (base_dir / orig_path).resolve()
        if orig_path_resolved.exists():
            target = images_dir / orig_path_resolved.name
            if not target.exists():
                shutil.move(str(orig_path_resolved), str(target))
            new_rel = images_dir.name + "/" + target.name
            return match.group(1) + new_rel + match.group(3)
        return match.group(0)
    return pattern.sub(repl, md_text)


def main() -> None:
    project_root = Path(__file__).resolve().parent
    # 注意：确保这里的路径正确指向你存放Excel的最外层文件夹
    # 根据你的描述，template_dir 应该是包含 "1-饼形图" 等子文件夹的目录
    template_dir = project_root / "template"
    
    output_root = project_root / "output"
    output_root.mkdir(parents=True, exist_ok=True)

    # --- 修改开始 ---
    # 使用 rglob 递归查找所有子文件夹下的 .xlsx 文件
    # 或者使用: glob("**/*.xlsx")
    xlsx_files = sorted(template_dir.rglob("*.xlsx"))
    # --- 修改结束 ---

    if not xlsx_files:
        # 如果这里报错，请打印 print(template_dir) 检查路径是否正确
        raise FileNotFoundError(f"No .xlsx file found in: {template_dir} (or its subdirectories)")

    converter = DocumentConverter()

    for source in xlsx_files:
        # 为了避免不同子文件夹下同名文件（如多个文件夹里都有 "1.xlsx") 冲突
        # 这里使用 source.stem 作为输出文件夹名，这样每个文件都有独立的输出目录
        out_dir = output_root / source.stem
        images_dir = out_dir / "images"
        out_dir.mkdir(parents=True, exist_ok=True)
        images_dir.mkdir(parents=True, exist_ok=True)

        # 将原始 Excel 文件也复制一份到输出目录
        shutil.copy2(source, out_dir / source.name)

        print(f"正在处理: {source.name} ...")
        
        try:
            result = converter.convert(str(source))

            md_path = out_dir / "index.md"
            result.document.save_as_markdown(md_path, image_mode=ImageRefMode.REFERENCED)

            md_text = md_path.read_text(encoding="utf-8")
            md_text_rewritten = _rewrite_image_paths(md_text, md_path.parent, images_dir)
            md_path.write_text(md_text_rewritten, encoding="utf-8")

            print(f"完成: {source.name} -> {out_dir}")
        except Exception as e:
            print(f"处理文件 {source} 时出错: {e}")


if __name__ == "__main__":
    main()