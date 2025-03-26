Office.onReady(() => {
  document.getElementById("sendPrompt").onclick = async () => {
    const prompt = document.getElementById("userPrompt").value.trim();
    const resultDiv = document.getElementById("result");

    if (!prompt) {
      resultDiv.innerText = "Please enter a description!";
      return;
    }

    resultDiv.innerText = "Generating formula...";

    const apiKey = "sk-571ba23044df464ebff3ecbed278f1b8"; // 🔑 Replace with your actual DeepSeek API key

    try {
      const response = await fetch("https://api.deepseek.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiKey}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "deepseek-chat", // or "deepseek-reasoner" based on your preference
          messages: [
            { "role": "system", "content": "You are an expert in Excel formulas. Provide only the formula without explanation." },
            { "role": "user", "content": prompt }
          ],
          max_tokens: 1000
        })
      });

      const data = await response.json();
      console.log("API Response:", data);

      if (data.choices && data.choices.length > 0) {
        const formula = data.choices[0].message.content.trim();
        resultDiv.innerText = `Generated Formula:\n${formula}`;

        // Insert formula into selected Excel cell
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.values = [[formula]];
          await context.sync();
        });
      } else {
        resultDiv.innerText = "Error: No valid response from DeepSeek.";
        console.error("DeepSeek Error:", data);
      }
    } catch (error) {
      console.error("Fetch Error:", error);
      resultDiv.innerText = "Error contacting DeepSeek API. Check console for details.";
    }
  };
});
console.log("API Response:", data);