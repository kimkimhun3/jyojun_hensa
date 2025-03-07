filePath = 'python.exe'
exec(`${pythonScriptPath} "${filePath}"`, (error, stdout, stderr) => {
  if (error) {
      console.error(`Error executing Python script: ${error}`);
      return;
  }
  console.log(`Python script output: ${stdout}`);
});