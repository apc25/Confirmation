import win32com.client as win32


class Hwp:
    # Visible so the file-access security prompt (shown when the automation module is not registered) is noticed.
    def __init__(self, visible=True):
        self.app = win32.Dispatch("HWPFrame.HwpObject")
        self.app.XHwpWindows.Item(0).Visible = visible
        self.app.SetMessageBoxMode(0x1)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.app.Clear(1)
        self.app.Quit()

    def open(self, path, fmt="HWP"):
        if not self.app.Open(str(path), fmt, "forceopen:true"):
            raise RuntimeError(f"한글에서 열지 못함: {path}")

    def append(self, path):
        """Insert another HWP at the end on a new page, hiding the header/footer inherited from the last page."""
        act = self.app.HAction
        self.app.MovePos(3, 0, 0)
        ins = self.app.HParameterSet.HInsertFile
        act.GetDefault("InsertFile", ins.HSet)
        ins.FileName = str(path)
        ins.KeepSection = 1
        ins.KeepCharshape = 1
        ins.KeepParashape = 1
        ins.KeepStyle = 1
        if not act.Execute("InsertFile", ins.HSet):
            raise RuntimeError(f"문서 끼워 넣기 실패: {path}")
        self.app.MovePos(3, 0, 0)
        hide = self.app.HParameterSet.HPageHiding
        act.GetDefault("PageHiding", hide.HSet)
        hide.Fields = 1 | 2 | 16
        act.Execute("PageHiding", hide.HSet)

    def close_doc(self):
        self.app.Clear(1)

    def text(self):
        """Body text including tables; headers/footers are not included."""
        return self.app.GetTextFile("UNICODE", "")

    def full_text(self):
        """HWPML export: markup, but also covers headers/footers."""
        return self.app.GetTextFile("HWPML2X", "")

    def top(self):
        self.app.MovePos(2, 0, 0)

    def find(self, s):
        """Move the cursor forward to the next occurrence of s."""
        act, pset = self.app.HAction, self.app.HParameterSet.HFindReplace
        act.GetDefault("ForwardFind", pset.HSet)
        pset.FindString = s
        pset.IgnoreMessage = 1
        pset.Direction = self.app.FindDir("Forward")
        found = act.Execute("ForwardFind", pset.HSet)
        self.app.Run("Cancel")
        if not found:
            raise RuntimeError(f"찾지 못함: {s!r}")

    def run(self, *actions):
        for a in actions:
            self.app.Run(a)

    def insert_text(self, s):
        act, pset = self.app.HAction, self.app.HParameterSet.HInsertText
        act.GetDefault("InsertText", pset.HSet)
        pset.Text = s
        act.Execute("InsertText", pset.HSet)

    def _all_replace(self, find, repl):
        act, pset = self.app.HAction, self.app.HParameterSet.HFindReplace
        act.GetDefault("AllReplace", pset.HSet)
        pset.FindString = find
        pset.ReplaceString = repl
        pset.FindRegExp = 0
        pset.IgnoreMessage = 1
        pset.Direction = self.app.FindDir("AllDoc")
        act.Execute("AllReplace", pset.HSet)

    def replace_everywhere(self, find, repl):
        """Replace in headers/footers too; verified on the HWPML export."""
        n = self.full_text().count(find)
        if n:
            self._all_replace(find, repl)
            if find in self.full_text().replace(repl, ""):
                raise RuntimeError(f"머리말·꼬리말 바꾸기 실패: {find!r}")
        return n

    def save(self, path, fmt):
        path.parent.mkdir(parents=True, exist_ok=True)
        if not self.app.SaveAs(str(path), fmt, ""):
            raise RuntimeError(f"저장 실패({fmt}): {path}")

    def replace(self, find, repl):
        """Replace every occurrence and return how many were replaced; verified by re-reading the text."""
        before = self.text()
        n = before.count(find)
        if n == 0:
            return 0
        self._all_replace(find, repl)
        after = self.text()
        ok = find not in after if not repl else after.count(repl) - before.count(repl) == n
        if not ok:
            raise RuntimeError(f"바꾸기 결과 불일치: {find!r} {n}곳")
        return n
