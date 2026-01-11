import { useEffect, useMemo, useRef, useState } from "react";
import "./App.css";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { saveAs } from "file-saver";

function tokenizeWithSpaces(text) {
  return text.match(/\S+|\s+/g) || [];
}
function isSpaceToken(token) {
  return /^\s+$/.test(token);
}

function shouldCapitalizeAt(text, cursorPos) {
  if (cursorPos === 0) return true;

  let i = cursorPos - 1;
  while (i >= 0 && /\s/.test(text[i])) i--;
  while (i >= 0 && /["'’”)\]]/.test(text[i])) i--;

  if (i < 0) return true;
  return text[i] === "." || text[i] === "!" || text[i] === "?";
}

function isSentenceBoundaryToken(tok) {
  // newline counts as boundary too
  if (/\n/.test(tok)) return true;
  return /[.!?]/.test(tok);
}

export default function App() {
  const [sourceText, setSourceText] = useState("");
  const [typedText, setTypedText] = useState("");

  const [isSourceLocked, setIsSourceLocked] = useState(false);
  const [isSessionActive, setIsSessionActive] = useState(false);

  const [autoScroll, setAutoScroll] = useState(true);
  const [forcedWordIndex, setForcedWordIndex] = useState(null);

  // Focus Mode
  const [focusMode, setFocusMode] = useState(false);

  // Privacy Mode (OPTIONAL): blur source until hover
  const [privacyMode, setPrivacyMode] = useState(false);

  // Reset behavior
  const [resetArmed, setResetArmed] = useState(false);
  const [showResetAllModal, setShowResetAllModal] = useState(false);

  const typingRef = useRef(null);
  const currentWordRef = useRef(null);

  useEffect(() => {
    if (isSessionActive) typingRef.current?.focus();
  }, [isSessionActive]);

  // Esc toggles focus mode (only during an active session)
  useEffect(() => {
    function onKeyDown(e) {
      if (e.key === "Escape" && isSessionActive && !showResetAllModal) {
        setFocusMode((v) => !v);
      }
    }
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [isSessionActive, showResetAllModal]);

  function handleLockSource() {
    setIsSourceLocked(true);
  }

  function handleStartSession() {
    if (!isSourceLocked) return alert("Lock the source first.");
    if (!sourceText.trim()) return alert("Paste some source text first.");

    setIsSessionActive(true);

    // Start in focus mode
    setFocusMode(true);

    // Starting cancels any "armed reset"
    setResetArmed(false);
  }

  function partialReset() {
    setTypedText("");
    setIsSessionActive(false);
    setForcedWordIndex(null);
    setFocusMode(false);
  }

  function fullResetAll() {
    setTypedText("");
    setIsSessionActive(false);
    setForcedWordIndex(null);

    setSourceText("");
    setIsSourceLocked(false);

    setFocusMode(false);
  }

  function handleResetClick() {
    if (showResetAllModal) return;

    if (!resetArmed) {
      partialReset();
      setResetArmed(true);
      return;
    }

    setShowResetAllModal(true);
  }

  async function handleDownloadDocx() {
    if (!typedText.trim()) {
      alert("Nothing to download yet. Type something first.");
      return;
    }

    const lines = typedText.split("\n");
    const doc = new Document({
      sections: [
        {
          properties: {},
          children: lines.map(
            (line) => new Paragraph({ children: [new TextRun(line)] })
          ),
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, "typed-text.docx");
  }

  // Tokens
  const sourceTokens = useMemo(() => tokenizeWithSpaces(sourceText), [sourceText]);
  const typedTokens = useMemo(() => tokenizeWithSpaces(typedText), [typedText]);

  // Word token indices: wordIndex -> tokenIndex
  const wordTokenIndices = useMemo(() => {
    const idxs = [];
    for (let i = 0; i < sourceTokens.length; i++) {
      if (!isSpaceToken(sourceTokens[i])) idxs.push(i);
    }
    return idxs;
  }, [sourceTokens]);

  // Strict matching memo
  const { autoMatchedWordCount, currentWordIndex, isCurrentWordWrong } = useMemo(() => {
    if (!isSessionActive) {
      return { autoMatchedWordCount: 0, currentWordIndex: 0, isCurrentWordWrong: false };
    }

    const sourceWords = sourceTokens.filter((t) => !isSpaceToken(t));
    const typedWords = typedTokens.filter((t) => !isSpaceToken(t));

    let i = 0;
    const max = Math.min(sourceWords.length, typedWords.length);
    for (; i < max; i++) {
      if (sourceWords[i] !== typedWords[i]) break;
    }

    const autoCount = i;
    const targetIndex = forcedWordIndex !== null ? forcedWordIndex : autoCount;

    let wrong = false;
    const typedHasWordHere = typedWords.length > targetIndex;

    if (typedHasWordHere) {
      const expected = sourceWords[targetIndex] ?? "";
      const actual = typedWords[targetIndex] ?? "";

      const movedToNextWord = typedWords.length > targetIndex + 1;
      const lastChar = typedText.slice(-1);
      const endsWithSeparator = /[\s.,!?;:)\]"'’”]/.test(lastChar);

      const committed = movedToNextWord || endsWithSeparator;
      if (committed) wrong = actual !== expected;
    }

    return {
      autoMatchedWordCount: autoCount,
      currentWordIndex: targetIndex,
      isCurrentWordWrong: wrong,
    };
  }, [isSessionActive, sourceTokens, typedTokens, typedText, forcedWordIndex]);

  // Sentence range
  const sentenceRange = useMemo(() => {
    if (!isSessionActive) return null;
    if (wordTokenIndices.length === 0) return null;

    const safeWordIndex = Math.min(currentWordIndex, wordTokenIndices.length - 1);
    const currentTokIndex = wordTokenIndices[safeWordIndex];

    let start = 0;
    for (let i = currentTokIndex - 1; i >= 0; i--) {
      if (isSentenceBoundaryToken(sourceTokens[i])) {
        start = i + 1;
        break;
      }
    }

    let end = sourceTokens.length - 1;
    for (let i = currentTokIndex; i < sourceTokens.length; i++) {
      if (isSentenceBoundaryToken(sourceTokens[i])) {
        end = i;
        break;
      }
    }

    return { start, end };
  }, [isSessionActive, sourceTokens, wordTokenIndices, currentWordIndex]);

  // Progress
  const sourceWordsCount = useMemo(() => {
    return sourceTokens.filter((t) => !isSpaceToken(t)).length;
  }, [sourceTokens]);

  const progressPercent = useMemo(() => {
    if (!sourceWordsCount) return 0;
    const pct = Math.floor((autoMatchedWordCount / sourceWordsCount) * 100);
    return Math.max(0, Math.min(100, pct));
  }, [autoMatchedWordCount, sourceWordsCount]);

  // Release forced mode when auto catches up
  useEffect(() => {
    if (!isSessionActive) return;
    if (forcedWordIndex === null) return;
    if (autoMatchedWordCount >= forcedWordIndex) setForcedWordIndex(null);
  }, [isSessionActive, forcedWordIndex, autoMatchedWordCount]);

  // Auto-scroll
  useEffect(() => {
    if (!isSessionActive) return;
    if (!autoScroll) return;

    currentWordRef.current?.scrollIntoView({
      block: "center",
      inline: "nearest",
      behavior: "smooth",
    });
  }, [isSessionActive, autoScroll, currentWordIndex, isCurrentWordWrong]);

  // Ctrl + Enter: jump forward 1 word
  useEffect(() => {
    if (!isSessionActive) return;

    function handleKeyDown(e) {
      if (e.ctrlKey && e.key === "Enter") {
        e.preventDefault();
        setForcedWordIndex((prev) =>
          prev === null ? currentWordIndex + 1 : prev + 1
        );
      }
    }

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isSessionActive, currentWordIndex]);

  // Render source with sentence highlight + word highlight
  const renderedSource = useMemo(() => {
    if (!isSourceLocked && !isSessionActive) return null;

    let wordCounter = 0;

    return sourceTokens.map((tok, idx) => {
      const inSentence =
        isSessionActive &&
        sentenceRange &&
        idx >= sentenceRange.start &&
        idx <= sentenceRange.end;

      if (isSpaceToken(tok)) {
  const baseClass = inSentence ? "sSentenceSpace" : "sSpace";
  const blurClass = privacyMode ? "pBlur" : "";
  return (
    <span key={idx} className={`${baseClass} ${blurClass}`}>
      {tok}
    </span>
  );
}


      const wordIndex = wordCounter;
      wordCounter++;

      let cls = "sRest";

      if (isSessionActive) {
        if (wordIndex < currentWordIndex) {
          cls = "sMatched";
        } else if (wordIndex === currentWordIndex) {
          cls = isCurrentWordWrong ? "sWrong" : "sNext";
        } else if (inSentence) {
          cls = "sSentence";
        } else {
          cls = "sRest";
        }
      }

      const isCurrent = isSessionActive && wordIndex === currentWordIndex;

      return (
        <span
          key={idx}
          className={`${cls} ${privacyMode && !isCurrent ? "pBlur" : ""}`}
          ref={isCurrent ? currentWordRef : null}
          onClick={() => {
            if (!isSessionActive) return;
            setForcedWordIndex(wordIndex);
          }}
          style={{ cursor: isSessionActive ? "pointer" : "default" }}
          title={isSessionActive ? "Click to resync here" : ""}
        >
          {tok}
        </span>
      );
    });
  }, [
  isSourceLocked,
  isSessionActive,
  sourceTokens,
  currentWordIndex,
  isCurrentWordWrong,
  sentenceRange,
  privacyMode,
]);


  return (
    <div className={`app ${showResetAllModal ? "blurred" : ""}`}>
      <header className="topbar">
        <div>
          <div className="title">Typing more easy</div>
        </div>

        {/* NORMAL MODE HEADER */}
        {!focusMode && (
          <div className="actions">
            <label className="toggle">
              <input
                type="checkbox"
                checked={autoScroll}
                onChange={(e) => setAutoScroll(e.target.checked)}
                disabled={!isSessionActive}
              />
              <span>Auto-scroll</span>
            </label>

            <label className="toggle">
              <input
                type="checkbox"
                checked={privacyMode}
                onChange={(e) => setPrivacyMode(e.target.checked)}
              />
              <span>Privacy</span>
            </label>

            <button
              className="btn"
              onClick={handleLockSource}
              disabled={isSourceLocked || isSessionActive || !sourceText.trim()}
            >
              {isSourceLocked ? "Source Locked" : "Lock Source"}
            </button>

            <button
              className="btn primary"
              onClick={handleStartSession}
              disabled={!isSourceLocked || isSessionActive}
            >
              {isSessionActive ? "Session Running" : "Start Session"}
            </button>

            <button
              className="btn"
              onClick={handleResetClick}
              disabled={!isSourceLocked && !sourceText.trim()}
              title={
                resetArmed
                  ? "Click again to reset EVERYTHING (will ask confirmation)"
                  : "Reset typing/session (keeps source). Click again to reset all."
              }
            >
              Reset
            </button>
          </div>
        )}

        {/* FOCUS MODE HEADER */}
        {focusMode && (
          <div className="actions">
            <button
              className="btn"
              onClick={() => setFocusMode(false)}
              title="Exit focus (Esc)"
            >
              Exit focus
            </button>

            <label className="toggle" title="Blur source until hover">
              <input
                type="checkbox"
                checked={privacyMode}
                onChange={(e) => setPrivacyMode(e.target.checked)}
              />
              <span>Privacy</span>
            </label>

            <button
              className="btn"
              onClick={handleResetClick}
              disabled={!isSourceLocked && !sourceText.trim()}
              title={
                resetArmed
                  ? "Click again to reset EVERYTHING (will ask confirmation)"
                  : "Reset typing/session (keeps source). Click again to reset all."
              }
            >
              Reset
            </button>

            <button className="btn" onClick={handleDownloadDocx} disabled={!typedText.trim()}>
              Download .docx
            </button>
          </div>
        )}
      </header>

      {isSessionActive && (
        <div className="progressWrap" aria-label="Progress">
          <div className="progressRow">
            <div className="progressLabel">{progressPercent}%</div>
            <div className="progressTrack">
              <div className="progressFill" style={{ width: `${progressPercent}%` }} />
            </div>
          </div>
        </div>
      )}

      <main className="panes">
        {/* LEFT */}
        <section className="pane">
          <div className="paneHeader">
            <span className="paneTitle">Source</span>
            {!focusMode && (
              <span className="pill">{isSourceLocked ? "Locked" : "Editable"}</span>
            )}
          </div>

          {!isSourceLocked && !isSessionActive ? (
            <textarea
              className="textarea"
              value={sourceText}
              onChange={(e) => setSourceText(e.target.value)}
              placeholder="Paste your study text here..."
              spellCheck={false}
            />
          ) : (
            <div
              className={`sourceViewer noCopy ${privacyMode ? "privacyOn" : ""}`}
              onCopy={(e) => e.preventDefault()}
              onCut={(e) => e.preventDefault()}
              onContextMenu={(e) => e.preventDefault()}
              onMouseDown={(e) => {
                if (e.shiftKey) e.preventDefault();
              }}
            >
              {renderedSource}
            </div>
          )}
        </section>

        {/* RIGHT */}
        <section className="pane">
          <div className="paneHeader">
            <span className="paneTitle">Typing</span>

            {!focusMode && (
              <div className="paneHeaderRight">
                <button className="btn" onClick={handleDownloadDocx} disabled={!typedText.trim()}>
                  Download .docx
                </button>
                <span className="pill">{isSessionActive ? "Active" : "Waiting"}</span>
              </div>
            )}
          </div>

          <textarea
            ref={typingRef}
            className="textarea"
            value={typedText}
            onChange={(e) => setTypedText(e.target.value)}
            onKeyDown={(e) => {
              if (!isSessionActive) return;

              // Word-like auto-capitalization
              if (e.ctrlKey || e.metaKey || e.altKey) return;
              if (e.key.length !== 1) return;
              if (!/[a-zA-Z]/.test(e.key)) return;

              const el = e.currentTarget;
              const start = el.selectionStart ?? 0;
              const end = el.selectionEnd ?? 0;
              if (start !== end) return;

              const needCap = shouldCapitalizeAt(typedText, start);
              if (!needCap) return;
              if (e.key === e.key.toUpperCase()) return;

              e.preventDefault();
              const before = typedText.slice(0, start);
              const after = typedText.slice(end);
              setTypedText(before + e.key.toUpperCase() + after);

              requestAnimationFrame(() => {
                typingRef.current?.setSelectionRange(start + 1, start + 1);
              });
            }}
            placeholder={
              isSessionActive ? "Type here..." : "Click Start Session (after locking source)"
            }
            spellCheck={false}
            disabled={!isSessionActive}
          />
        </section>
      </main>

      {/* Reset-all modal */}
      {showResetAllModal && (
        <div className="modalOverlay" role="dialog" aria-modal="true">
          <div className="modalCard">
            <div className="modalTitle">Reset everything?</div>
            <div className="modalText">Are you sure you want to reset it all?</div>

            <div className="modalActions">
              <button
                className="btn"
                onClick={() => {
                  setShowResetAllModal(false);
                  setResetArmed(false);
                }}
              >
                No
              </button>

              <button
                className="btn primary"
                onClick={() => {
                  fullResetAll();
                  setShowResetAllModal(false);
                  setResetArmed(false);
                }}
              >
                Yes
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
