const toBase64Clean = (blob) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onloadend = () => {
    const result = reader.result;
    if (!result) return resolve('');
    resolve(result.substring(result.indexOf(',') + 1).trim());
  };
  reader.onerror = reject;
  reader.readAsDataURL(blob);
});

export const uploadMoodleResultToSharePoint = async (txtBlob, pdfBlob, xlsxBlob, folderName, txtName, pdfName, xlsxName, flowUrl) => {
  try {
    const [txtContent, pdfContent, xlsxContent] = await Promise.all([
      toBase64Clean(txtBlob),
      toBase64Clean(pdfBlob),
      xlsxBlob ? toBase64Clean(xlsxBlob) : Promise.resolve(''),
    ]);

    const response = await fetch(flowUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ folderName, csvName: txtName, pdfName, xlsxName, csvContent: txtContent, pdfContent, xlsxContent }),
    });

    if (response.ok) {
      console.log('Moodle-Upload erfolgreich! Ordner:', folderName);
      return true;
    }
    console.error('Fehler beim Moodle-Upload', await response.text());
    return false;
  } catch (error) {
    console.error('Netzwerkfehler:', error);
    return false;
  }
};

export const uploadToSharePoint = async (csvBlob, pdfBlob, xlsxBlob, folderName, csvName, pdfName, xlsxName, flowUrl) => {
  try {
    const [csvContent, pdfContent, xlsxContent] = await Promise.all([
      toBase64Clean(csvBlob),
      toBase64Clean(pdfBlob),
      xlsxBlob ? toBase64Clean(xlsxBlob) : Promise.resolve(''),
    ]);

    const response = await fetch(flowUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ folderName, csvName, pdfName, xlsxName, csvContent, pdfContent, xlsxContent }),
    });

    if (response.ok) {
      console.log('Upload erfolgreich! Ordner:', folderName);
      return true;
    }
    console.error('Fehler beim Upload', await response.text());
    return false;
  } catch (error) {
    console.error('Netzwerkfehler:', error);
    return false;
  }
};
