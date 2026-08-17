let winCount = 0;
let lossCount = 0;

const winLossChart = new Chart(document.getElementById("winLossChart"), {
  type: 'pie',
  data: {
    labels: ['Winning', 'Losing'],
    datasets: [{
      data: [0, 0],
      backgroundColor: ['#00c853', '#d50000']
    }]
  }
});

const profitChart = new Chart(document.getElementById("profitChart"), {
  type: 'bar',
  data: {
    labels: [],
    datasets: [{
      label: 'Profit',
      data: [],
      backgroundColor: '#c9a34a'
    }]
  }
});

document.getElementById("tradeForm").addEventListener("submit", function(e) {
  e.preventDefault();

  const symbol = document.getElementById("symbol").value;
  const option = document.getElementById("option").value;
  const strike = document.getElementById("strike").value;
  const buyPrice = parseFloat(document.getElementById("buy").value);
  const sellPrice = parseFloat(document.getElementById("sell").value);
  const status = document.getElementById("status").value;

  const profit = (sellPrice - buyPrice).toFixed(2);
  const percent = ((profit / buyPrice) * 100).toFixed(0);

  const row = document.createElement("tr");
  row.className = status === "win" ? "win" : "loss";

  row.innerHTML = `
    <td>${symbol}</td>
    <td>${option}</td>
    <td>${strike}</td>
    <td>$${buyPrice.toFixed(2)}</td>
    <td>$${sellPrice.toFixed(2)}</td>
    <td>${profit > 0 ? "+" : ""}$${profit}</td>
    <td>${percent > 0 ? "+" : ""}${percent}%</td>
    <td>${status === "win" ? "✅" : "❌"}</td>
  `;

  document.getElementById("tradeTableBody").appendChild(row);

  if (status === "win") winCount++;
  else lossCount++;

  winLossChart.data.datasets[0].data = [winCount, lossCount];
  winLossChart.update();

  profitChart.data.labels.push(symbol);
  profitChart.data.datasets[0].data.push(profit);
  profitChart.update();

  this.reset();
});
