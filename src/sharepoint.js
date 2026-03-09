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

export const uploadToSharePoint = async (csvBlob, pdfBlob, folderName, csvName, pdfName, flowUrl) => {
  try {
    const [csvContent, pdfContent] = await Promise.all([
      toBase64Clean(csvBlob),
      toBase64Clean(pdfBlob),
    ]);

    const response = await fetch(flowUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ folderName, csvName, pdfName, csvContent, pdfContent }),
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
