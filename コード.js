// Webアプリのエントリーポイント
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('People Core')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// 従業員データを取得する関数
function getEmployeeData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('従業員');
    
    if (!sheet) {
      throw new Error('従業員シートが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow < 2) {
      return [];
    }
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // 必要な列のインデックスを取得
    const employeeIdIndex = headers.indexOf('社員番号');
    const nameIndex = headers.indexOf('名前');
    const employmentTypeIndex = headers.indexOf('雇用形態');
    const activeStatusIndex = headers.indexOf('在籍中フラグ');
    
    // データを取得
    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    
    // 必要な項目のみを抽出して配列に変換
    const employees = data.map(row => {
      return {
        '社員番号': row[employeeIdIndex] || '',
        '名前': row[nameIndex] || '',
        '雇用形態': row[employmentTypeIndex] || '',
        '在籍中フラグ': row[activeStatusIndex] || ''
      };
    });
    
    return employees;
    
  } catch (error) {
    console.error('getEmployeeDataでエラーが発生しました:', error.toString());
    throw error;
  }
}

// 特定の従業員の詳細データを取得
function getEmployeeDetail(employeeId) {
  const employees = getEmployeeData();
  const employee = employees.find(emp => emp['社員番号'] === employeeId);
  
  if (!employee) {
    throw new Error('指定された社員が見つかりません');
  }
  
  return employee;
}

// フィルタ用のユニークな値を取得する関数
function getUniqueValues(columnName) {
  const employees = getEmployeeData();
  const values = new Set();
  
  employees.forEach(employee => {
    if (employee[columnName]) {
      values.add(employee[columnName]);
    }
  });
  
  return Array.from(values).sort();
}

// デバッグ用：サンプルデータの確認
function debugGetSampleData() {
  console.log('=== サンプルデータ取得テスト開始 ===');
  
  try {
    const employees = getEmployeeData();
    
    if (employees.length === 0) {
      console.log('従業員データが取得できませんでした');
      return 'データなし';
    }
    
    console.log('\n取得された従業員数:', employees.length);
    
    // 最初の3件のデータを表示
    console.log('\n最初の3件のデータ:');
    for (let i = 0; i < Math.min(3, employees.length); i++) {
      console.log(`\n--- 従業員 ${i + 1} ---`);
      const employee = employees[i];
      Object.keys(employee).forEach(key => {
        console.log(`${key}: ${employee[key]}`);
      });
    }
    
    // データの統計情報
    console.log('\n=== データ統計情報 ===');
    const columnStats = {};
    Object.keys(employees[0]).forEach(key => {
      const nonEmptyCount = employees.filter(emp => emp[key] && emp[key] !== '').length;
      columnStats[key] = {
        total: employees.length,
        nonEmpty: nonEmptyCount,
        emptyRate: ((employees.length - nonEmptyCount) / employees.length * 100).toFixed(1) + '%'
      };
    });
    
    console.log('\n各列の入力状況:');
    Object.keys(columnStats).forEach(key => {
      const stat = columnStats[key];
      console.log(`${key}: ${stat.nonEmpty}/${stat.total} (空白率: ${stat.emptyRate})`);
    });
    
    console.log('\n=== サンプルデータ取得テスト完了 ===');
    return `${employees.length}件のデータを取得しました`;
    
  } catch (error) {
    console.error('エラーが発生しました:', error.toString());
    return 'エラー: ' + error.toString();
  }
}
