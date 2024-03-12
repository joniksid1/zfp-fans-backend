const generateUniqueFileName = () => {
  const timestamp = new Date().getTime();
  const randomBytes = crypto.randomBytes(16).toString('hex');
  return `newFile_${timestamp}_${randomBytes}.xlsx`;
};

module.exports = { generateUniqueFileName };
