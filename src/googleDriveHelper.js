/**
 * Helper utilities for integrating with Google Drive API v3 on the client side.
 */

// Generate a random 16-character string preserving the original file extension
export function generateRandom16CharFilename(originalName) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < 16; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  // Try to extract extension, default to empty
  const dotIndex = originalName.lastIndexOf('.');
  const extension = dotIndex !== -1 ? originalName.substring(dotIndex) : '';
  return result + extension;
}

// Find a folder in the user's Drive or create it if it doesn't exist
async function getOrCreateFolder(accessToken, folderName) {
  const query = `name = '${folderName}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}`;

  try {
    const searchResponse = await fetch(searchUrl, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!searchResponse.ok) {
      throw new Error(`Failed to search folder: ${searchResponse.statusText}`);
    }

    const searchData = await searchResponse.json();
    if (searchData.files && searchData.files.length > 0) {
      // Folder exists, return its ID
      return searchData.files[0].id;
    }

    // Folder doesn't exist, create it
    const createUrl = 'https://www.googleapis.com/drive/v3/files';
    const metadata = {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
    };

    const createResponse = await fetch(createUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(metadata),
    });

    if (!createResponse.ok) {
      throw new Error(`Failed to create folder: ${createResponse.statusText}`);
    }

    const folderData = await createResponse.json();
    return folderData.id;
  } catch (error) {
    console.error('Error in getOrCreateFolder:', error);
    throw error;
  }
}

/**
 * Uploads a file to the user's Google Drive.
 * @param {string} accessToken Google OAuth2 access token.
 * @param {File} file The file object from input.
 * @param {string} customName Pre-generated 16-character filename.
 * @param {string} folderName Target folder name.
 * @returns {Promise<{id: string, name: string}>} The uploaded file ID and name.
 */
export async function uploadFileToDrive(accessToken, file, customName, folderName = 'SIPA FTK UnHar Files') {
  try {
    const folderId = await getOrCreateFolder(accessToken, folderName);

    const metadata = {
      name: customName,
      parents: [folderId],
    };

    // Construct a multipart upload request
    const boundary = '-------314159265358979323846';
    const delimiter = `\r\n--${boundary}\r\n`;
    const closeDelimiter = `\r\n--${boundary}--`;

    const reader = new FileReader();

    // Read the file as binary string to construct multipart request
    const fileDataPromise = new Promise((resolve, reject) => {
      reader.onload = (event) => resolve(event.target.result);
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });

    const fileContentArrayBuffer = await fileDataPromise;
    const metadataString = JSON.stringify(metadata);

    // Build the request body parts
    const header = `${delimiter}Content-Type: application/json; charset=UTF-8\r\n\r\n${metadataString}${delimiter}Content-Type: ${file.type || 'application/octet-stream'}\r\n\r\n`;
    
    // Convert header string to Uint8Array
    const encoder = new TextEncoder();
    const headerUint8 = encoder.encode(header);
    const footerUint8 = encoder.encode(closeDelimiter);
    
    // Combine parts into a single Uint8Array
    const bodyUint8 = new Uint8Array(headerUint8.length + fileContentArrayBuffer.byteLength + footerUint8.length);
    bodyUint8.set(headerUint8, 0);
    bodyUint8.set(new Uint8Array(fileContentArrayBuffer), headerUint8.length);
    bodyUint8.set(footerUint8, headerUint8.length + fileContentArrayBuffer.byteLength);

    const uploadUrl = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
    const response = await fetch(uploadUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `multipart/related; boundary=${boundary}`,
      },
      body: bodyUint8,
    });

    if (!response.ok) {
      throw new Error(`Upload failed: ${response.statusText}`);
    }

    const data = await response.json();
    return {
      id: data.id,
      name: data.name,
    };
  } catch (error) {
    console.error('Error in uploadFileToDrive:', error);
    throw error;
  }
}
