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

    def _criteria(self, data, tags, level):
        blocks = []
        if data.get("tags"):
            if not any(t in data["tags"] for t in tags):
                return blocks
        blocks.append(f"{'#' * level} {data['title']}")

        if data.get("level").upper() == "MUST":
            blocks.append("<span class='level level-must'>Muss-Kriterium</span>")
        elif data.get("level").upper() == "SHOULD":
            blocks.append("<span class='level level-should'>Soll-Kriterium</span>")
        elif data.get("level").upper() == "MAY":
            blocks.append("<span class='level level-may'>Kann-Kriterium</span>")
        for b in data['description'].split("\n\n"):
            blocks.append(b)
        if data.get("comment"):
            blocks.append('!!! note "Kommentar"')
            blocks.append('')
            blocks.append(indent(data['comment'], "    "))
        if data.get("sub"):
            for l in data["sub"]:
                blocks += self._criteria(l, tags, level + 1)
        return blocks

    def _blocks(self, file, tags):
        with open("data/modules/" + file + ".yml", "r") as f:
            data = yaml.safe_load(f.read())

        blocks = []
        for l in data:
            blocks += self._criteria(l, tags, 2)


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
