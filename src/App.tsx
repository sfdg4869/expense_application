import { useCallback, useEffect, useState } from "react";
import type { Defaults, TransactionRow } from "./types";

export default function App() {
  const [statementPath, setStatementPath] = useState("");
  const [templatePath, setTemplatePath] = useState("");
  const [imagePaths, setImagePaths] = useState<string[]>([]);
  const [transactions, setTransactions] = useState<TransactionRow[]>([]);
  const [defaults, setDefaults] = useState<Defaults>({
    account: "복리후생비",
    client: "엑셈",
    userName: "정경수",
    detail: "야근식대",
  });
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<{ type: "ok" | "err"; text: string } | null>(null);

  useEffect(() => {
    window.expenseApi.loadDefaults().then(setDefaults).catch(() => {});
  }, []);

  const loadStatement = useCallback(async (path: string) => {
    const rows = await window.expenseApi.parseStatement(path);
    setTransactions(rows);
    setMessage({ type: "ok", text: `거래 ${rows.length}건을 불러왔습니다.` });
  }, []);

  const pickStatement = async () => {
    const p = await window.expenseApi.pickFile([
      { name: "Excel", extensions: ["xls", "xlsx"] },
    ]);
    if (!p) return;
    setStatementPath(p);
    try {
      await loadStatement(p);
    } catch (e) {
      setMessage({ type: "err", text: String(e) });
    }
  };

  const pickTemplate = async () => {
    const p = await window.expenseApi.pickFile([
      { name: "Excel", extensions: ["xls", "xlsx"] },
    ]);
    if (p) setTemplatePath(p);
  };

  
  const pickImageFolder = async () => {
    const folder = await window.expenseApi.pickFolder();
    if (!folder) return;
    const files = await window.expenseApi.loadImagesFromFolder(folder);
    setImagePaths(files);
    setMessage({ type: "ok", text: `영수증 ${files.length}장 (폴더)` });
  };

  const pickImages = async () => {
    const files = await window.expenseApi.pickImages();
    if (files.length) {
      setImagePaths(files);
      setMessage({ type: "ok", text: `영수증 ${files.length}장 선택됨` });
    }
  };

  const updatePersonal = (index: number, value: number) => {
    setTransactions((prev) =>
      prev.map((t, i) => (i === index ? { ...t, personalUseAmount: value } : t))
    );
  };

  const saveDefaults = async () => {
    await window.expenseApi.saveDefaults(defaults);
    setMessage({ type: "ok", text: "기본값을 저장했습니다." });
  };

  const generate = async () => {
    setMessage(null);
    if (!statementPath || !templatePath) {
      setMessage({ type: "err", text: "명세서와 양식 파일을 모두 선택하세요." });
      return;
    }

    let rows = transactions;
    if (rows.length === 0) {
      try {
        rows = await window.expenseApi.parseStatement(statementPath);
        setTransactions(rows);
      } catch (e) {
        setMessage({
          type: "err",
          text: `명세서를 읽지 못했습니다.
${String(e)}

카드사 명세서(두 번째 이미지 형식)를 선택했는지 확인하세요.`,
        });
        return;
      }
    }
    if (rows.length === 0) {
      setMessage({ type: "err", text: "거래 내역이 없습니다. 위에서 「카드 명세서」를 다시 눌러 「거래 N건」이 보이는지 확인하세요." });
      return;
    }

    const dir = templatePath.replace(/\\[^\\]+$/, "");
    const monthMatch = rows[0]?.useDate?.match(/\.(\d{2})\./);
    const month = monthMatch ? monthMatch[1] : "00";
    const defaultName = `(카드)제경비신청서_${month}월_완성.xls`;
    const out =
      (await window.expenseApi.pickSave(defaultName)) ??
      `${dir}\\${defaultName}`;

    setBusy(true);
    try {
      const result = await window.expenseApi.generate({
        statementPath,
        templatePath,
        outputPath: out,
        transactions: rows,
        defaults,
        imagePaths,
      });
      if (result.success) {
        setMessage({
          type: "ok",
          text: `완성 파일이 저장되었습니다.\n${result.outputPath ?? out}`,
        });
      } else {
        setMessage({ type: "err", text: result.error ?? "생성 실패" });
      }
    } catch (e) {
      setMessage({ type: "err", text: String(e) });
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="app">
      <h1>경비신청서 자동 작성</h1>
      <p className="sub">
        카드 명세서(14열) → 회사 양식 「이용내역 붙여넣기」 영역에 채움 + 영수증 첨부
      </p>

      <div className="card">
        <h2>1. 파일 선택</h2>
        <div className="row">
          <button type="button" onClick={pickStatement}>
            카드 명세서
          </button>
          <span className="path">{statementPath || "미선택"}{transactions.length > 0 ? ` · 거래 ${transactions.length}건` : statementPath ? " · ⚠ 명세서 다시 선택" : ""}</span>
        </div>
        <div className="row">
          <button type="button" onClick={pickTemplate}>
            회사 양식
          </button>
          <span className="path">{templatePath || "미선택"}</span>
        </div>
        <div className="row">
          <button type="button" onClick={pickImages}>
            영수증 선택
          </button>
          <button type="button" onClick={pickImageFolder}>
            영수증 폴더
          </button>
          <span className="path">
            {imagePaths.length ? `${imagePaths.length}장 선택됨` : "미선택 (선택 사항)"}
          </span>
        </div>
      </div>

      <div className="card">
        <h2>2. 기본 입력값</h2>
        <div className="grid">
          <div>
            <label>계정</label>
            <input
              value={defaults.account}
              onChange={(e) => setDefaults({ ...defaults, account: e.target.value })}
            />
          </div>
          <div>
            <label>거래처명/소재지</label>
            <input
              value={defaults.client}
              onChange={(e) => setDefaults({ ...defaults, client: e.target.value })}
            />
          </div>
          <div>
            <label>사용자명</label>
            <input
              value={defaults.userName}
              onChange={(e) => setDefaults({ ...defaults, userName: e.target.value })}
            />
          </div>
          <div>
            <label>업무 상세내용</label>
            <input
              value={defaults.detail}
              onChange={(e) => setDefaults({ ...defaults, detail: e.target.value })}
            />
          </div>
        </div>
        <div className="row" style={{ marginTop: 12 }}>
          <button type="button" onClick={saveDefaults}>
            기본값 저장
          </button>
        </div>
      </div>

      {transactions.length > 0 && (
        <div className="card">
          <h2>3. 거래 검토 (개인사용금액만 수정)</h2>
          <table>
            <thead>
              <tr>
                <th>이용일자</th>
                <th>가맹점</th>
                <th>청구금액</th>
                <th>개인사용</th>
                <th>경비신청금액</th>
              </tr>
            </thead>
            <tbody>
              {transactions.map((t, i) => (
                <tr key={i}>
                  <td>{t.useDate}</td>
                  <td>{t.merchant}</td>
                  <td>{t.domesticClaimAmount.toLocaleString()}</td>
                  <td>
                    <input
                      type="number"
                      min={0}
                      value={t.personalUseAmount}
                      onChange={(e) =>
                        updatePersonal(i, parseFloat(e.target.value) || 0)
                      }
                      style={{ width: 80 }}
                    />
                  </td>
                  <td>
                    {(t.domesticClaimAmount - t.personalUseAmount).toLocaleString()}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="card">
        <button
          type="button"
          className="primary"
          disabled={busy}
          onClick={generate}
        >
          {busy ? "생성 중…" : "완성 엑셀 만들기"}
        </button>
        <p className="sub" style={{ marginTop: 8, marginBottom: 0 }}>
          저장 위치를 고르지 않으면 양식과 같은 폴더에 자동 저장됩니다.
        </p>
      </div>

      {message && (
        <div className={`msg ${message.type}`} style={{ whiteSpace: "pre-wrap" }}>
          {message.text}
        </div>
      )}
    </div>
  );
}

