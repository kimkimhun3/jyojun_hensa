// Sample extracted data
const extractedData = [
    { timestamp: '2025/03/03 16:58:58', speed: '214.5' },
    { timestamp: '2025/03/03 16:58:59', speed: '2325.0' },
    { timestamp: '2025/03/03 16:59:00', speed: '2611.9' }
  ];
  
  // Extract only the speed values
const speeds = extractedData.map(item => item.speed);

console.log(speeds);

const range = Array.from({ length: speeds.length }, (_, i) => i);

console.log(range);