"""
export_vba.py
将 xlam 文件中的所有 VBA 模块导出为可读文本文件，供 git 追踪。

用法：
    python scripts/export_vba.py                     # 自动查找仓库根目录下的 .xlam
    python scripts/export_vba.py MyAddin.xlam        # 指定 xlam 路径
    python scripts/export_vba.py MyAddin.xlam src/vba  # 指定输出目录
"""

import sys
import os
import re
from pathlib import Path

try:
    from oletools.olevba import VBA_Parser, TYPE_MODULE, TYPE_CLASS, TYPE_FORM
except ImportError:
    print("❌ 缺少依赖，请运行：pip install oletools")
    sys.exit(1)


# VBA 模块类型 → 文件扩展名映射
EXT_MAP = {
    TYPE_MODULE: ".bas",   # 普通模块
    TYPE_CLASS:  ".cls",   # 类模块
    TYPE_FORM:   ".frm",   # 窗体模块
}
DEFAULT_EXT = ".bas"

# 导出时过滤掉的内置模块（这些是 Excel 自动生成的，没有自定义代码）
SKIP_MODULES = {"ThisWorkbook", "Sheet"}


def find_xlam(search_dir: Path) -> Path | None:
    """在指定目录（及子目录）中查找第一个 .xlam 文件"""
    for p in search_dir.rglob("*.xlam"):
        # 跳过备份文件
        if ".bak_" not in p.name:
            return p
    return None


def sanitize_filename(name: str) -> str:
    """去除文件名中的非法字符"""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def should_skip(module_name: str) -> bool:
    """判断是否跳过该模块"""
    for prefix in SKIP_MODULES:
        if module_name.startswith(prefix):
            return True
    return False


def export_vba(xlam_path: Path, output_dir: Path) -> list[str]:
    """
    导出 xlam 中所有 VBA 模块到 output_dir。
    返回已导出的文件路径列表。
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    exported = []

    parser = VBA_Parser(str(xlam_path))

    if not parser.detect_vba_macros():
        print(f"⚠️  {xlam_path.name} 中未检测到 VBA 代码")
        return exported

    # 记录已处理的模块名（避免重复）
    seen = set()

    for _, _, vba_filename, code in parser.extract_macros():
        # vba_filename 形如 "Module1.bas" 或 "Module1"
        raw_name = Path(vba_filename).stem
        ext      = Path(vba_filename).suffix or DEFAULT_EXT

        if raw_name in seen:
            continue
        seen.add(raw_name)

        if should_skip(raw_name):
            print(f"  ⏭  跳过内置模块：{raw_name}")
            continue

        # 统一扩展名（oletools 有时给出错误后缀）
        filename = sanitize_filename(raw_name) + ext
        out_path = output_dir / filename

        # 写入时统一使用 UTF-8，去除 Windows 行尾
        clean_code = code.replace("\r\n", "\n").replace("\r", "\n")
        out_path.write_text(clean_code, encoding="utf-8")

        print(f"  ✅ 已导出：{filename}  ({len(clean_code.splitlines())} 行)")
        exported.append(str(out_path))

    parser.close()
    return exported


def write_manifest(output_dir: Path, xlam_path: Path, exported: list[str]):
    """生成 manifest.txt，记录本次导出的元信息"""
    import datetime
    manifest_path = output_dir / "manifest.txt"
    lines = [
        f"# VBA 导出清单",
        f"# 源文件：{xlam_path}",
        f"# 导出时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"# 模块数量：{len(exported)}",
        "",
    ]
    for f in sorted(exported):
        lines.append(Path(f).name)
    manifest_path.write_text("\n".join(lines), encoding="utf-8")


def main():
    args = sys.argv[1:]

    # 确定 xlam 路径
    if args and args[0].endswith(".xlam"):
        xlam_path = Path(args[0])
        args = args[1:]
    else:
        # 自动从当前目录或仓库根目录查找
        repo_root = Path.cwd()
        xlam_path = find_xlam(repo_root)
        if xlam_path is None:
            print("❌ 未找到 .xlam 文件，请在仓库根目录运行或手动指定路径")
            sys.exit(1)
        print(f"🔍 自动检测到：{xlam_path}")

    if not xlam_path.exists():
        print(f"❌ 文件不存在：{xlam_path}")
        sys.exit(1)

    # 确定输出目录
    output_dir = Path(args[0]) if args else Path("src/vba")

    print(f"\n📦 正在从 {xlam_path.name} 导出 VBA 模块 → {output_dir}/\n")
    exported = export_vba(xlam_path, output_dir)

    if exported:
        write_manifest(output_dir, xlam_path, exported)
        print(f"\n🎉 共导出 {len(exported)} 个模块，清单已写入 {output_dir}/manifest.txt")
    else:
        print("\n⚠️  没有导出任何模块")

    # 返回导出的文件列表（供 hook 脚本调用）
    return exported


if __name__ == "__main__":
    main()
