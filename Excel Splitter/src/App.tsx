import React, { useState, useMemo, useRef } from 'react';
import { read, utils, write } from 'xlsx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { useVirtualizer } from '@tanstack/react-virtual';
import { FileDown, UploadCloud, Shuffle, Trash2, UserX } from 'lucide-react';
import './index.css';

export default function App() {
  const [data, setData] = useState<any[][]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [chunkSize, setChunkSize] = useState(100);
  const [phoneColIndex, setPhoneColIndex] = useState(0);
  const [nameColIndex, setNameColIndex] = useState(0);
  const [isDragActive, setIsDragActive] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [controlName, setControlName] = useState('');
  const [controlPhone, setControlPhone] = useState('');
  const [controlRows, setControlRows] = useState<{ name: string, phone: string }[]>(() => {
    const saved = localStorage.getItem('controlRows');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  React.useEffect(() => {
    localStorage.setItem('controlRows', JSON.stringify(controlRows));
  }, [controlRows]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      setFileName(file.name);
      processFile(file);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setFileName(file.name);
      processFile(file);
    }
    e.target.value = '';
  };

  function findPhoneColumn(rows: any[][]) {
    let bestColIndex = 0;
    let maxPhoneCount = 0;
    for (let col = 0; col < (rows[0]?.length || 0); col++) {
      let phoneCount = 0;
      for (let row = 0; row < Math.min(rows.length, 20); row++) {
        let val = String(rows[row]?.[col] ?? '').replace(/\D/g, '');
        if (val.length >= 8 && val.length <= 15) {
          phoneCount++;
        }
      }
      if (phoneCount > maxPhoneCount) {
        maxPhoneCount = phoneCount;
        bestColIndex = col;
      }
    }
    return bestColIndex;
  }

  function processFile(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = read(e.target?.result, { type: 'binary' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows: any[][] = utils.sheet_to_json(firstSheet, { header: 1 });

      let headerRowIndex = 0;
      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const row = rows[i];
        if (Array.isArray(row) && row.some(cell => {
          const val = String(cell ?? '').toLowerCase();
          return val.includes('nome') || val.includes('telefone') || val.includes('celular');
        })) {
          headerRowIndex = i;
          break;
        }
      }

      const parsedHeaders = rows[headerRowIndex] || [];
      const dataRows = rows.slice(headerRowIndex + 1).filter(row =>
        Array.isArray(row) && row.some(cell => cell != null && String(cell).trim() !== '')
      );

      const phoneCol = findPhoneColumn(dataRows);

      let nameCol = parsedHeaders.findIndex(h => String(h).toLowerCase().includes('nome'));
      if (nameCol === -1) {
        nameCol = phoneCol === 1 ? 0 : (phoneCol === 0 ? 1 : 0);
      }

      const fixedRows = dataRows.map(row => {
        const newRow = [...row];
        const phoneStr = String(newRow[phoneCol] ?? '').replace(/\D/g, '');
        const isPhoneValid = phoneStr.length >= 8 && phoneStr.length <= 15;

        if (!isPhoneValid) {
          for (let c = 0; c < Math.max(newRow.length, parsedHeaders.length, 2); c++) {
            if (c !== phoneCol) {
              const valStr = String(newRow[c] ?? '').trim();
              const valNums = valStr.replace(/\D/g, '');
              const valLetters = valStr.replace(/[0-9\s\-\+\(\)]/g, '');
              // Se tiver números suficientes e quase nenhuma letra, assumimos que é o telefone que "escorregou"
              if (valNums.length >= 8 && valNums.length <= 15 && valLetters.length < 3) {
                newRow[phoneCol] = newRow[c];
                newRow[c] = '';
                break;
              }
            }
          }
        }
        return newRow;
      });

      setHeaders(parsedHeaders);
      setData(fixedRows);
      setPhoneColIndex(phoneCol);
      setNameColIndex(nameCol);
    };
    reader.readAsBinaryString(file);
  }

  function removeDuplicates(rows: any[][]) {
    const bestRowIndices = new Map<string, number>();

    rows.forEach((row, idx) => {
      const phone = String(row[phoneColIndex] ?? '').trim();
      if (!phone) return;

      const name = String(row[nameColIndex] ?? '').trim();
      const existingIdx = bestRowIndices.get(phone);

      if (existingIdx === undefined) {
        bestRowIndices.set(phone, idx);
      } else {
        const existingName = String(rows[existingIdx][nameColIndex] ?? '').trim();
        if (!existingName && name) {
          bestRowIndices.set(phone, idx);
        }
      }
    });

    const validIndices = new Set(bestRowIndices.values());

    return rows.filter((row, idx) => {
      const phone = String(row[phoneColIndex] ?? '').trim();
      if (!phone) return true;
      return validIndices.has(idx);
    });
  }

  const duplicatesSet = useMemo(() => {
    const seen = new Set();
    const duplicates = new Set();
    data.forEach(row => {
      const val = String(row[phoneColIndex] ?? '').trim();
      if (val) {
        if (seen.has(val)) {
          duplicates.add(val);
        } else {
          seen.add(val);
        }
      }
    });
    return duplicates;
  }, [data, phoneColIndex]);

  const uniqueData = useMemo(() => removeDuplicates(data), [data, phoneColIndex]);
  const duplicateCount = data.length - uniqueData.length;

  function shuffleData() {
    setData(prev => {
      const newData = [...prev];
      for (let i = newData.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [newData[i], newData[j]] = [newData[j], newData[i]];
      }
      return newData;
    });
  }

  function removeEmptyNames() {
    setData(prev => prev.filter(row => {
      const nameVal = String(row[nameColIndex] ?? '').trim();
      const phoneVal = String(row[phoneColIndex] ?? '').replace(/\D/g, '');
      return nameVal !== '' && phoneVal.length >= 8;
    }));
  }

  const parentRef = useRef<HTMLDivElement>(null);

  const rowVirtualizer = useVirtualizer({
    count: data.length + controlRows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 48,
    overscan: 5,
  });

  const RowComponent = React.useCallback(({ index, style }: { index: number, style: React.CSSProperties }) => {
    if (index >= data.length) {
      const ctrlIndex = index - data.length;
      const ctrl = controlRows[ctrlIndex];
      const rowLen = Math.max(headers.length, nameColIndex + 1, phoneColIndex + 1);
      const row = new Array(rowLen).fill('');
      row[nameColIndex] = ctrl.name;
      row[phoneColIndex] = ctrl.phone.replace(/\D/g, '');

      return (
        <div
          style={{
            ...style,
            display: 'flex',
            background: '#e0f2fe',
            borderBottom: '1px solid var(--border)'
          }}
          className="table-row"
        >
          {row.map((cell, i) => (
            <div
              key={i}
              className="table-cell"
              style={{
                flex: i === phoneColIndex ? 2 : 1,
                color: 'var(--primary)',
                fontWeight: 500
              }}
            >
              {cell ?? ''}
            </div>
          ))}
        </div>
      );
    }

    const row = data[index];
    const val = String(row[phoneColIndex] ?? '').trim();
    const isDuplicate = duplicatesSet.has(val);

    return (
      <div
        style={{
          ...style,
          display: 'flex',
          background: isDuplicate ? 'var(--danger-bg)' : index % 2 === 0 ? '#fafafa' : '#ffffff',
          borderBottom: '1px solid var(--border)'
        }}
        className="table-row"
      >
        {row.map((cell, i) => (
          <div
            key={i}
            className="table-cell"
            style={{
              flex: i === phoneColIndex ? 2 : 1,
              color: isDuplicate ? 'var(--danger)' : 'var(--text)'
            }}
          >
            {cell ?? ''}
          </div>
        ))}
      </div>
    );
  }, [data, controlRows, headers, nameColIndex, phoneColIndex, duplicatesSet]);

  async function exportZip() {
    const rowsToExport = removeDuplicates(data);
    const zip = new JSZip();
    let fileNum = 1;

    for (let i = 0; i < rowsToExport.length; i += chunkSize) {
      const chunk = rowsToExport.slice(i, i + chunkSize);

      const chunkWithControls = [...chunk];
      controlRows.forEach(ctrl => {
        const rowLen = Math.max(headers.length, nameColIndex + 1, phoneColIndex + 1);
        const newRow = new Array(rowLen).fill('');
        newRow[nameColIndex] = ctrl.name;
        newRow[phoneColIndex] = ctrl.phone.replace(/\D/g, '');
        chunkWithControls.push(newRow);
      });

      const sheet = utils.aoa_to_sheet([headers, ...chunkWithControls]);
      const workbook = utils.book_new();
      utils.book_append_sheet(workbook, sheet, 'Lista');
      const excelBuffer = write(workbook, { type: 'array', bookType: 'xlsx' });
      zip.file(`lista_${String(fileNum).padStart(3, '0')}.xlsx`, excelBuffer);
      fileNum++;
    }

    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, 'listas.zip');
  }

  return (
    <div className="office-container">
      <h2 className="office-title">
        <FileDown style={{ width: 32, height: 32, color: 'var(--primary)' }} />
        Divisor de Excel
      </h2>

      <div className="office-card">
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`office-dropzone ${isDragActive ? 'active' : ''}`}
          onClick={() => document.getElementById('file-upload')?.click()}
        >
          <div style={{ pointerEvents: 'none' }}>
            <UploadCloud style={{ width: 48, height: 48, marginBottom: 16, color: isDragActive ? 'var(--primary)' : '#cbd5e1' }} />
            <p style={{ margin: 0, fontWeight: 500, fontSize: 16, color: isDragActive ? 'var(--primary)' : 'var(--text)' }}>
              {isDragActive ? 'Solte o arquivo aqui' : (fileName ? `Arquivo carregado: ${fileName}` : 'Clique ou arraste a planilha (.xlsx, .csv)')}
            </p>
          </div>
        </div>
        <input
          id="file-upload"
          type="file"
          accept=".xlsx, .xls, .csv"
          style={{ display: 'none' }}
          onChange={handleFileInput}
        />

        <div style={{ marginTop: 24, padding: 16, backgroundColor: '#f8fafc', borderRadius: 8, border: '1px solid var(--border)' }}>
          <h3 style={{ margin: '0 0 12px 0', fontSize: 14, color: 'var(--text)' }}>Contatos Fixos (Em azul adicionados ao final de cada lote)</h3>
          <div style={{ display: 'flex', gap: 8, marginBottom: controlRows.length > 0 ? 12 : 0 }}>
            <input
              type="text"
              placeholder="Nome"
              value={controlName}
              onChange={e => setControlName(e.target.value.replace(/[^a-zA-ZáàãâéèêíïóôõöúçñÁÀÃÂÉÈÊÍÏÓÔÕÖÚÇÑ]/g, ''))}
              style={{ flex: 1, padding: '8px 12px', borderRadius: 6, border: '1px solid var(--border)', outline: 'none' }}
            />
            <input
              type="text"
              placeholder="Telefone +55 (xx) xxxxx-xxxx"
              value={controlPhone}
              onChange={e => {
                let raw = e.target.value;
                
                if (raw === '+55 ' || raw === '+55' || raw === '+5' || raw === '+') {
                  setControlPhone('');
                  return;
                }

                let digits = raw.replace(/\D/g, '');
                
                if (digits.startsWith('55') && digits.length > 2) {
                  digits = digits.substring(2);
                }
                
                if (digits.length > 11) digits = digits.slice(0, 11);

                let formatted = '';
                if (digits.length > 0) {
                  formatted = '+55 ';
                  if (digits.length > 10) {
                    formatted += `(${digits.slice(0, 2)}) ${digits.slice(2, 7)}-${digits.slice(7)}`;
                  } else if (digits.length > 6) {
                    formatted += `(${digits.slice(0, 2)}) ${digits.slice(2, 6)}-${digits.slice(6)}`;
                  } else if (digits.length > 2) {
                    formatted += `(${digits.slice(0, 2)}) ${digits.slice(2)}`;
                  } else {
                    formatted += `(${digits}`;
                  }
                }
                setControlPhone(formatted);
              }}
              style={{ flex: 1, padding: '8px 12px', borderRadius: 6, border: '1px solid var(--border)', outline: 'none' }}
            />
            <button
              className="btn btn-secondary"
              onClick={() => {
                if (controlName.trim() || controlPhone.trim()) {
                  const phoneOnly = controlPhone.replace(/\D/g, '');
                  
                  setData(prev => prev.filter(row => {
                    const rowPhone = String(row[phoneColIndex] ?? '').replace(/\D/g, '');
                    return rowPhone !== phoneOnly;
                  }));
                  
                  setControlRows(prev => [...prev, { name: controlName, phone: controlPhone }]);
                  setControlName('');
                  setControlPhone('');
                }
              }}
            >
              Adicionar
            </button>
          </div>
          {controlRows.length > 0 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {controlRows.map((row, i) => (
                <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#fff', padding: '8px 12px', borderRadius: 4, border: '1px solid var(--border)' }}>
                  <span style={{ fontSize: 13, fontWeight: 500 }}>{row.name || '-'} <span style={{ color: 'var(--text-light)', marginLeft: 8 }}>{row.phone || '-'}</span></span>
                  <button
                    style={{ background: 'none', border: 'none', color: 'var(--danger)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                    onClick={() => setControlRows(prev => prev.filter((_, idx) => idx !== i))}
                  >
                    <Trash2 style={{ width: 16, height: 16 }} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="controls-row">
          <div className="input-group">
            <label>Usuários por arquivo:</label>
            <input
              type="number"
              value={chunkSize}
              onChange={(e) => setChunkSize(Number(e.target.value))}
            />
          </div>
          <button onClick={exportZip} className="btn btn-primary" disabled={data.length === 0}>
            <FileDown style={{ width: 18, height: 18 }} />
            Exportar ZIP
          </button>
        </div>
      </div>

      <div className="stats-bar">
        <div className="stat-group">
          <div className="stat-item">
            <span className="stat-label">Total de Linhas</span>
            <span className="stat-value">{data.length}</span>
          </div>
          <div className="stat-item">
            <span className="stat-label">Duplicados</span>
            <span className="stat-value" style={{ color: duplicateCount > 0 ? 'var(--danger)' : 'var(--text)' }}>
              {duplicateCount}
            </span>
          </div>
        </div>
        <div className="actions-group">
          <button onClick={shuffleData} disabled={data.length === 0} className="btn btn-secondary">
            <Shuffle style={{ width: 16, height: 16 }} />
            Embaralhar
          </button>
          <button onClick={removeEmptyNames} disabled={data.length === 0} className="btn btn-secondary">
            <UserX style={{ width: 16, height: 16 }} />
            Remover Vazios
          </button>
          <button onClick={() => setData(uniqueData)} disabled={duplicateCount === 0} className="btn btn-danger">
            <Trash2 style={{ width: 16, height: 16 }} />
            Remover Duplicados
          </button>
        </div>
      </div>

      <div>
        <div className="table-header">
          {headers.map((header, i) => (
            <div key={i} style={{ flex: i === phoneColIndex ? 2 : 1 }}>
              {header || `Coluna ${i + 1}`}
            </div>
          ))}
        </div>
        <div className="table-container" style={{ height: 500 }}>
          {data.length > 0 && (
            <div ref={parentRef} style={{ height: '100%', overflow: 'auto' }}>
              <div style={{ height: rowVirtualizer.getTotalSize(), position: 'relative' }}>
                {rowVirtualizer.getVirtualItems().map((virtualRow) => (
                  <RowComponent key={virtualRow.index} index={virtualRow.index} style={{ position: 'absolute', top: virtualRow.start, height: virtualRow.size, width: '100%' }} />
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
