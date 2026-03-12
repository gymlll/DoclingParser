from __future__ import annotations

import argparse
from pathlib import Path

from llm import LLMAPIClient


def pick_markdown(path_arg: str | None) -> Path:
    if path_arg:
        p = Path(path_arg).expanduser().resolve()
        if not p.exists():
            raise FileNotFoundError(f"Markdown file not found: {p}")
        return p

    output_dir = Path(__file__).resolve().parent / "output"
    md_files = sorted(output_dir.glob("*.md"), key=lambda x: x.stat().st_mtime, reverse=True)
    if not md_files:
        raise FileNotFoundError(f"No markdown file found in: {output_dir}")
    return md_files[0]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Chat with parsed markdown using llm.py client.")
    parser.add_argument("--file", type=str, default=None, help="Markdown file path. Default: latest .md in ./output")
    parser.add_argument("--max-chars", type=int, default=30000, help="Max chars from markdown to inject into context")
    return parser


def main() -> None:
    args = build_parser().parse_args()

    md_path = pick_markdown(args.file)
    text = md_path.read_text(encoding="utf-8", errors="replace")
    if len(text) > args.max_chars:
        text = text[: args.max_chars]

    client = LLMAPIClient()

    system_prompt = (
        "你是一个严谨的文档问答助手。"
        "你只能基于已提供的文档内容回答；若文档中没有明确答案，请明确说明。"
        f"\n\n以下是用户上传并授权你分析的文档内容（文件: {md_path.name}）:\n\n{text}"
    )
    messages = [{"role": "system", "content": system_prompt}]

    print(f"Loaded markdown: {md_path}")
    print(f"Model: {client.model_name}")
    print("Enter your question. Type /exit to quit.")

    while True:
        q = input("\nYou> ").strip()
        if not q:
            continue
        if q.lower() in {"/exit", "exit", "quit", "/quit"}:
            print("Bye.")
            break

        messages.append({"role": "user", "content": q})
        request_messages = messages + [
            {
                "role": "user",
                "content": (
                    "If the previous user message is in Chinese, answer in Chinese, "
                    "but first interpret the question accurately from the document context."
                ),
            }
        ]
        print("Assistant> ", end="", flush=True)
        try:
            response = client.chat_completion(messages=request_messages, stream=True, temperature=0.2, max_tokens=600)
            answer_parts = []
            for chunk in response:
                delta = chunk.choices[0].delta.content
                if delta:
                    answer_parts.append(delta)
                    print(delta, end="", flush=True)
            print()
            answer = "".join(answer_parts).strip()
            if answer:
                messages.append({"role": "assistant", "content": answer})
            else:
                messages.pop()
        except Exception as e:
            messages.pop()
            print(f"\nAssistant> API call failed: {e}")


if __name__ == "__main__":
    main()
