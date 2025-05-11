// App.js
import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

// קומפוננטה לטבלה עם אפשרות מיון
const SortableTable = ({ data, columns, title, formatNumber = (num) => num }) => {
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // מיון הנתונים
  const requestSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  // סידור הנתונים לפי המיון הנבחר
  const sortedData = React.useMemo(() => {
    let sortableItems = [...data];
    if (sortConfig.key) {
      sortableItems.sort((a, b) => {
        if (a[sortConfig.key] === null || a[sortConfig.key] === undefined) return 1;
        if (b[sortConfig.key] === null || b[sortConfig.key] === undefined) return -1;
        
        const aValue = typeof a[sortConfig.key] === 'string' ? a[sortConfig.key].toLowerCase() : a[sortConfig.key];
        const bValue = typeof b[sortConfig.key] === 'string' ? b[sortConfig.key].toLowerCase() : b[sortConfig.key];
        
        if (aValue < bValue) {
          return sortConfig.direction === 'asc' ? -1 : 1;
        }
        if (aValue > bValue) {
          return sortConfig.direction === 'asc' ? 1 : -1;
        }
        return 0;
      });
    }
    return sortableItems;
  }, [data, sortConfig]);

  // סמן לכיוון המיון
  const getSortDirectionIndicator = (key) => {
    if (sortConfig.key !== key) return null;
    return sortConfig.direction === 'asc' ? ' ▲' : ' ▼';
  };

  return (
    <div className="table-container">
      <h2 className="table-title">{title}</h2>
      <div className="table-wrapper">
        <table className="data-table">
          <thead>
            <tr>
              {columns.map(column => (
                <th 
                  key={column.key}
                  className="table-header"
                  onClick={() => column.sortable !== false && requestSort(column.key)}
                >
                  {column.label}
                  {column.sortable !== false && getSortDirectionIndicator(column.key)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sortedData.map((item, index) => (
              <tr key={index} className={index % 2 === 0 ? 'row-even' : 'row-odd'}>
                {columns.map(column => (
                  <td key={`${index}-${column.key}`} className="table-cell">
                    {column.render 
                      ? column.render(item[column.key], item) 
                      : column.isNumeric && item[column.key] !== null
                        ? formatNumber(item[column.key]) + (column.unit || '')
                        : item[column.key] || '-'}
                  </td>
                ))}
              </tr>
            ))}
            {sortedData.length === 0 && (
              <tr>
                <td colSpan={columns.length} className="table-cell empty-message">
                  לא נמצאו נתונים
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const FinancialDashboard = () => {
  const [activeTab, setActiveTab] = useState('participation');
  const [data, setData] = useState({
    participation: { byYear: {}, institutions: [] },
    construction: { byYear: {}, projects: [] },
    loans: { byYear: {}, loansList: [] }
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [fileUploaded, setFileUploaded] = useState(false);

  // פונקציה לטעינת הקובץ
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        processExcelData(workbook);
        setFileUploaded(true);
      } catch (err) {
        console.error("Error processing file:", err);
        setError("אירעה שגיאה בעיבוד הקובץ: " + err.message);
      } finally {
        setLoading(false);
      }
    };
    
    reader.onerror = (event) => {
      setError("אירעה שגיאה בקריאת הקובץ");
      setLoading(false);
    };
    
    reader.readAsArrayBuffer(file);
  };

  // פונקציה לעיבוד נתוני האקסל
  const processExcelData = (workbook) => {
    // Parse participation data
    const participationSheet = workbook.Sheets["השתתפות מרכז בשוטף"];
    if (!participationSheet) {
      setError("גיליון 'השתתפות מרכז בשוטף' לא נמצא בקובץ");
      return;
    }

    const participationData = XLSX.utils.sheet_to_json(participationSheet, {
      header: 1,
      defval: null,
      blankrows: false
    });
    
    // Parse construction data
    const constructionSheet = workbook.Sheets["בינוי "];
    if (!constructionSheet) {
      setError("גיליון 'בינוי ' לא נמצא בקובץ");
      return;
    }

    const constructionData = XLSX.utils.sheet_to_json(constructionSheet, {
      header: 1,
      defval: null,
      blankrows: false
    });
    
    // Parse loans data
    const loanSheet = workbook.Sheets["הלוואת בעלות"];
    if (!loanSheet) {
      setError("גיליון 'הלוואת בעלות' לא נמצא בקובץ");
      return;
    }

    const loanData = XLSX.utils.sheet_to_json(loanSheet, {
      header: 1,
      defval: null,
      blankrows: false
    });

    // Process participation data
    const participationByYear = {};
    const institutions = [];
    
    participationData.forEach((row, index) => {
      if (index === 0) return; // Skip header
      
      const year = row[0];
      if (!year || typeof year !== 'string' || !year.includes("תשפ")) return;
      
      if (!participationByYear[year]) {
        participationByYear[year] = {
          agreementSum: 0,
          approvedSum: 0,
          finalSum: 0,
          count: 0
        };
      }
      
      participationByYear[year].count++;
      
      if (row[2] && !isNaN(row[2])) participationByYear[year].agreementSum += row[2];
      if (row[3] && !isNaN(row[3])) participationByYear[year].approvedSum += row[3];
      if (row[4] && !isNaN(row[4])) participationByYear[year].finalSum += row[4];
      
      institutions.push({
        year: year,
        name: row[1],
        agreementParticipation: row[2],
        approvedParticipation: row[3],
        finalParticipation: row[4],
        notes: row[5]
      });
    });
    
    // Process construction data
    const constructionByYear = {};
    const projects = [];
    
    constructionData.forEach((row, index) => {
      if (index === 0) return; // Skip header
      
      const year = row[0];
      if (!year || typeof year !== 'string' || !year.includes("תשפ")) return;
      
      if (!constructionByYear[year]) {
        constructionByYear[year] = {
          totalCost: 0,
          initialParticipation: 0,
          currentParticipation: 0,
          count: 0
        };
      }
      
      constructionByYear[year].count++;
      
      if (row[3] && !isNaN(row[3])) constructionByYear[year].totalCost += row[3];
      if (row[4] && !isNaN(row[4])) constructionByYear[year].initialParticipation += row[4];
      if (row[5] && !isNaN(row[5])) constructionByYear[year].currentParticipation += row[5];
      
      projects.push({
        year: year,
        institution: row[1],
        projectType: row[2],
        totalCost: row[3],
        initialParticipation: row[4],
        currentParticipation: row[5],
        notes: row[6]
      });
    });
    
    // Process loans data
    const loansByYear = {};
    const loansList = [];
    
    // Find header row
    let headerRowIndex = loanData.findIndex(row => row[0] === "שנה");
    
    if (headerRowIndex >= 0) {
      for (let i = headerRowIndex + 1; i < loanData.length; i++) {
        const row = loanData[i];
        if (!row[0]) continue;
        
        const year = row[0];
        if (!year || typeof year !== 'string' || !year.includes("תשפ")) continue;
        
        if (!loansByYear[year]) {
          loansByYear[year] = {
            totalAmount: 0,
            count: 0
          };
        }
        
        loansByYear[year].count++;
        if (row[3] && !isNaN(row[3])) loansByYear[year].totalAmount += row[3];
        
        loansList.push({
          year: year,
          institution: row[1],
          loanType: row[2],
          amount: row[3],
          repayment: row[4]
        });
      }
    }

    setData({
      participation: { 
        byYear: participationByYear, 
        institutions 
      },
      construction: { 
        byYear: constructionByYear, 
        projects 
      },
      loans: { 
        byYear: loansByYear, 
        loansList 
      }
    });
  };

  // פונקציה לעיצוב מספרים
  const formatNumber = (num) => {
    if (num === null || num === undefined) return '-';
    return new Intl.NumberFormat('he-IL').format(num);
  };

  return (
    <div className="dashboard-container">
      <h1 className="dashboard-title">דשבורד נתונים פיננסיים</h1>
      
      {/* טעינת קובץ */}
      <div className="file-upload-section">
        <label htmlFor="excel-file" className="file-upload-label">
          בחר/י קובץ אקסל:
          <input
            type="file"
            id="excel-file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="file-input"
          />
        </label>
        <div className="file-status">
          {loading && <p className="loading-message">טוען נתונים...</p>}
          {error && <p className="error-message">{error}</p>}
          {fileUploaded && <p className="success-message">הקובץ נטען בהצלחה!</p>}
        </div>
      </div>
      
      {fileUploaded && (
        <>
          {/* Tabs */}
          <div className="tabs-container">
            <button 
              className={`tab-button ${activeTab === 'participation' ? 'active-tab' : ''}`}
              onClick={() => setActiveTab('participation')}
            >
              השתתפות מרכז בשוטף
            </button>
            <button 
              className={`tab-button ${activeTab === 'construction' ? 'active-tab' : ''}`}
              onClick={() => setActiveTab('construction')}
            >
              בינוי
            </button>
            <button 
              className={`tab-button ${activeTab === 'loans' ? 'active-tab' : ''}`}
              onClick={() => setActiveTab('loans')}
            >
              הלוואות בעלות
            </button>
          </div>
          
          {/* סיכום לפי סוג נתונים */}
          {activeTab === 'participation' && (
            <div className="tab-content">
              {/* טבלת סיכום */}
              <SortableTable
                title="סיכום השתתפות לפי שנים"
                data={Object.keys(data.participation.byYear).map(year => ({
                  year,
                  count: data.participation.byYear[year].count,
                  agreementSum: data.participation.byYear[year].agreementSum,
                  approvedSum: data.participation.byYear[year].approvedSum,
                  finalSum: data.participation.byYear[year].finalSum
                }))}
                formatNumber={formatNumber}
                columns={[
                  { key: 'year', label: 'שנה', sortable: true },
                  { key: 'count', label: 'מספר מוסדות', sortable: true, isNumeric: true },
                  { 
                    key: 'agreementSum', 
                    label: 'השתתפות ע"פ הסכם', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                  },
                  { 
                    key: 'approvedSum', 
                    label: 'השתתפות מאושרת', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'  
                  },
                  { 
                    key: 'finalSum', 
                    label: 'השתתפות סופית', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'  
                  }
                ]}
              />
              
              {/* טבלת פירוט לפי שנים */}
              {Object.keys(data.participation.byYear).map(year => (
                <SortableTable
                  key={year}
                  title={`${year} - פירוט השתתפות לפי מוסדות`}
                  data={data.participation.institutions.filter(inst => inst.year === year)}
                  formatNumber={formatNumber}
                  columns={[
                    { key: 'name', label: 'שם מוסד', sortable: true },
                    { 
                      key: 'agreementParticipation', 
                      label: 'השתתפות ע"פ הסכם', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { 
                      key: 'approvedParticipation', 
                      label: 'השתתפות מאושרת', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { 
                      key: 'finalParticipation', 
                      label: 'השתתפות סופית', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { key: 'notes', label: 'הערות', sortable: true }
                  ]}
                />
              ))}
            </div>
          )}
          
          {/* תצוגת נתוני בינוי */}
          {activeTab === 'construction' && (
            <div className="tab-content">
              {/* טבלת סיכום */}
              <SortableTable
                title="סיכום פרויקטי בינוי לפי שנים"
                data={Object.keys(data.construction.byYear).map(year => ({
                  year,
                  count: data.construction.byYear[year].count,
                  totalCost: data.construction.byYear[year].totalCost,
                  initialParticipation: data.construction.byYear[year].initialParticipation,
                  currentParticipation: data.construction.byYear[year].currentParticipation
                }))}
                formatNumber={formatNumber}
                columns={[
                  { key: 'year', label: 'שנה', sortable: true },
                  { key: 'count', label: 'מספר פרויקטים', sortable: true, isNumeric: true },
                  { 
                    key: 'totalCost', 
                    label: 'סה"כ עלות פרויקטים', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                  },
                  { 
                    key: 'initialParticipation', 
                    label: 'השתתפות רשת ראשונית', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'  
                  },
                  { 
                    key: 'currentParticipation', 
                    label: 'השתתפות רשת נכון ל-1.5.25', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'  
                  }
                ]}
              />
              
              {/* טבלת פירוט לפי שנים */}
              {Object.keys(data.construction.byYear).map(year => (
                <SortableTable
                  key={year}
                  title={`${year} - פירוט פרויקטי בינוי`}
                  data={data.construction.projects.filter(proj => proj.year === year)}
                  formatNumber={formatNumber}
                  columns={[
                    { key: 'institution', label: 'שם מוסד', sortable: true },
                    { key: 'projectType', label: 'סוג פרויקט', sortable: true },
                    { 
                      key: 'totalCost', 
                      label: 'עלות פרויקט', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { 
                      key: 'initialParticipation', 
                      label: 'השתתפות ראשונית', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { 
                      key: 'currentParticipation', 
                      label: 'השתתפות נכון ל-1.5.25', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { key: 'notes', label: 'הערות', sortable: true }
                  ]}
                />
              ))}
            </div>
          )}
          
          {/* תצוגת נתוני הלוואות */}
          {activeTab === 'loans' && (
            <div className="tab-content">
              {/* טבלת סיכום */}
              <SortableTable
                title="סיכום הלוואות בעלות לפי שנים"
                data={Object.keys(data.loans.byYear).map(year => ({
                  year,
                  count: data.loans.byYear[year].count,
                  totalAmount: data.loans.byYear[year].totalAmount
                }))}
                formatNumber={formatNumber}
                columns={[
                  { key: 'year', label: 'שנה', sortable: true },
                  { key: 'count', label: 'מספר הלוואות', sortable: true, isNumeric: true },
                  { 
                    key: 'totalAmount', 
                    label: 'סה"כ סכום הלוואות', 
                    sortable: true, 
                    isNumeric: true,
                    render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                  }
                ]}
              />
              
              {/* טבלת פירוט לפי שנים */}
              {Object.keys(data.loans.byYear).map(year => (
                <SortableTable
                  key={year}
                  title={`${year} - פירוט הלוואות בעלות`}
                  data={data.loans.loansList.filter(loan => loan.year === year)}
                  formatNumber={formatNumber}
                  columns={[
                    { key: 'institution', label: 'שם מוסד', sortable: true },
                    { key: 'loanType', label: 'סוג הלוואה', sortable: true },
                    { 
                      key: 'amount', 
                      label: 'סכום הלוואה שנתי', 
                      sortable: true,
                      isNumeric: true,
                      render: (value) => value ? formatNumber(value) + ' ₪' : '-'
                    },
                    { key: 'repayment', label: 'החזר', sortable: true }
                  ]}
                />
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default FinancialDashboard;