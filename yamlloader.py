from textwrap import indent
from typing import MutableSequence
import xml.etree.ElementTree as etree

import yaml
from markdown import Markdown
from markdown.blockprocessors import BlockProcessor
from markdown.preprocessors import Preprocessor
from markdown.extensions import Extension
import re


class YamlDocProcessor(BlockProcessor):
    regex = re.compile(r"^::: ?(?P<file>[^ ]+)(?P<tags>( [^ ]+)*) *$", flags=re.MULTILINE)

    def __init__(self, md: Markdown):
        super().__init__(parser=md.parser)
        self.md = md

    def test(self, parent, block: str) -> bool:  # noqa: ARG002
        return bool(self.regex.search(block))

    def run(self, parent, blocks: MutableSequence[str]):
        block = blocks[0]
        match = self.regex.search(block)

        if match:
            e = etree.SubElement(parent, 'div')
            self.parser.parseBlocks(e, self._blocks(match.group("file"), match.group("tags").strip().split(" ")))
            blocks.pop(0)

    def _blocks(self, file, tags):
        with open("data/modules/" + file + ".yml", "r") as f:
            data = yaml.safe_load(f.read())

        blocks = []
        for l in data:
            if l.get("tags"):
                if not any(t in l["tags"] for t in tags):
                    continue
            blocks.append(f"## {l['title']}")
            if l.get("level") == "MUST":
                blocks.append("<span class='level level-must'>Muss-Kriterium</span>")
            elif l.get("level") == "SHOULD":
                blocks.append("<span class='level level-should'>Soll-Kriterium</span>")
            elif l.get("level") == "MAY":
                blocks.append("<span class='level level-may'>Kann-Kriterium</span>")
            blocks.append(l['description'])
            if l.get("comment"):
                blocks.append('!!! note "Kommentar"')
                blocks.append('')
                blocks.append(indent(l['comment'], "    "))


        return blocks


class YamlLoadExtension(Extension):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def extendMarkdown(self, md):
        md.registerExtension(self)

        md.parser.blockprocessors.register(
            YamlDocProcessor(md),
            "yamlload",
            priority=75,
        )


def makeExtension(**kwargs):  # pragma: no cover
    return YamlLoadExtension(**kwargs)
