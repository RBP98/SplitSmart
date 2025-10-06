function chooseFile() {
    document.getElementById("fileInput").click();
}

document.getElementById("fileInput").addEventListener("change", function() {
    var file = this.files[0];
    if (file) {
        document.getElementById("fileName").innerText = "Selected file: " + file.name;
    } else {
        document.getElementById("fileName").innerText = "";
    }
});
function downloadSpreadsheet() {
    fetch("/downloadInvoice", {
        method: "POST",
        // body: formData
    })
        .then(response => response.blob())
        .then(blob => {
            // Create a blob URL for the downloaded file
            const url = window.URL.createObjectURL(blob);

            // Create a link element
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = "abcd.xlsx";

            // Add the link to the document body
            document.body.appendChild(a);

            // Simulate a click on the link to trigger the download
            a.click();

            // Cleanup
            a.remove();
            window.URL.revokeObjectURL(url);
        })
        .catch(error => {
            console.error("Error:", error);
        });
}

document.getElementById("submitForm").addEventListener("submit", function(event) {
    event.preventDefault();
    var file = document.getElementById("fileInput").files[0];
    if (file) {
        var formData = new FormData();
        formData.append("file", file);

        fetch("http://localhost:8080/upload", {
            method: "POST",
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                renderInvoice(data);
            })
            .catch(error => {
                console.error("Error:", error);
            });
    } else {
        alert("Please select a file before submitting.");
    }
});
function renderInvoice(orderInfo) {
    const orderTable = document.getElementById("order-info");
    orderTable.innerHTML = `
        <p><strong>Order Number:</strong> ${orderInfo.orderNumber}</p>
        <p><strong>Date:</strong> ${new Date(orderInfo.date).toLocaleDateString()}</p>
        <table>
            <tr>
                <th>Item</th>
                <th>Amount</th>
                <th>Qty</th>
                <th>Status</th>
            </tr>
            ${orderInfo.itemList.map(item => `
                <tr>
                    <td class="item-name">${item.name}</td>
                    <td>$${item.amount.toFixed(2)}</td>
                    <td>${item.qty}</td>
                    <td>${item.status}</td>
                </tr>
            `).join('')}
            <tr class="total-row">
                <td colspan="4"><strong>Subtotal:</strong> $${orderInfo.subTotal.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Savings:</strong> $${orderInfo.savings.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Bag Fee:</strong> $${orderInfo.bagFee.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Tax:</strong> $${orderInfo.tax.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Driver Tip:</strong> $${orderInfo.driverTip.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Donation:</strong> $${orderInfo.donation.toFixed(2)}</td>
            </tr>
            <tr class="total-row">
                <td colspan="4"><strong>Total:</strong> $${orderInfo.total.toFixed(2)}</td>
            </tr>
        </table>
        <p><strong>Card Ending In:</strong> ${orderInfo.cardEndingIn}</p>
    `;
}
