from __future__ import annotations

from PySide6.QtCore import QObject, Signal


class HistoryCommand:
    def redo(self) -> None:
        raise NotImplementedError

    def undo(self) -> None:
        raise NotImplementedError


class HistoryManager(QObject):
    canUndoChanged = Signal(bool)
    canRedoChanged = Signal(bool)
    cleanChanged = Signal(bool)
    limitChanged = Signal(int)
    stateChanged = Signal()

    def __init__(self, limit: int = 100, parent=None) -> None:
        super().__init__(parent)
        self._commands: list[HistoryCommand] = []
        self._index = 0
        self._clean_index = 0
        self._limit = max(1, int(limit))
        self._is_executing = False

    @property
    def limit(self) -> int:
        return self._limit

    @property
    def is_executing(self) -> bool:
        return self._is_executing

    def can_undo(self) -> bool:
        return self._index > 0

    def can_redo(self) -> bool:
        return self._index < len(self._commands)

    def is_clean(self) -> bool:
        return self._clean_index >= 0 and self._index == self._clean_index

    def clear(self) -> None:
        self._commands.clear()
        self._index = 0
        self._clean_index = 0
        self._emit_state()

    def set_clean(self) -> None:
        self._clean_index = self._index
        self._emit_state()

    def set_limit(self, limit: int) -> None:
        normalized_limit = max(1, int(limit))
        if normalized_limit == self._limit:
            return
        self._limit = normalized_limit
        self._enforce_limit()
        self.limitChanged.emit(self._limit)
        self._emit_state()

    def push(self, command: HistoryCommand) -> bool:
        if command is None:
            return False

        if self._index < len(self._commands):
            del self._commands[self._index :]
            if self._clean_index > self._index:
                self._clean_index = -1

        self._is_executing = True
        try:
            command.redo()
        finally:
            self._is_executing = False

        self._commands.append(command)
        self._index += 1
        self._enforce_limit()
        self._emit_state()
        return True

    def undo(self) -> bool:
        if not self.can_undo():
            return False

        self._index -= 1
        self._is_executing = True
        try:
            self._commands[self._index].undo()
        finally:
            self._is_executing = False
        self._emit_state()
        return True

    def redo(self) -> bool:
        if not self.can_redo():
            return False

        self._is_executing = True
        try:
            self._commands[self._index].redo()
        finally:
            self._is_executing = False
        self._index += 1
        self._emit_state()
        return True

    def _enforce_limit(self) -> None:
        overflow = len(self._commands) - self._limit
        if overflow <= 0:
            return

        del self._commands[:overflow]
        self._index = max(0, self._index - overflow)
        if self._clean_index >= 0:
            if self._clean_index < overflow:
                self._clean_index = -1
            else:
                self._clean_index -= overflow

    def _emit_state(self) -> None:
        self.canUndoChanged.emit(self.can_undo())
        self.canRedoChanged.emit(self.can_redo())
        self.cleanChanged.emit(self.is_clean())
        self.stateChanged.emit()
