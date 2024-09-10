const Button = document.getElementById("Button");
const EmailInput = document.getElementById("email");
const PasswordInput = document.getElementById("password");

// > If the input is not null
EmailInput.addEventListener("change", () => {
  if (EmailInput.value.length > 0) {
    Button.classList.remove("Disable");
    Button.classList.add("Enable");

    // When clicking the button, send the token to the destination email
    Button.addEventListener("click", async () => {
      Button.style.backgroundColor = "#595959";

      const Email = document.getElementById("email").value;

      if (Email && Button.textContent.includes("Token")) {
        await SendingInformation(Email);

      } else if (Email && Button.textContent.includes("Login")) {
        await SendingLoginInformation(Email, PasswordInput.value);
      } else {
        console.error("Email is required");
      }

      // >  Set timeout to return its initial state
      setTimeout(() => {
        Button.style.backgroundColor = "#000";
      }, 200);
    });
  } else {
    Button.classList.add("Disable");
    Button.classList.remove("Enable");
  }
});

async function SendingInformation(Email) {
  const destination = "/token";

  const response = await fetch(destination, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-cache",
      Pragma: "no-cache",
    },
    body: JSON.stringify({ gmail: Email }),
  });

  if (!response.ok) {
    throw new Error("Network response was not ok");
  } else {
    // > If the connection is ok
    if (response.status == 200) {
      // > Remove the attribute `disabled` to enable password input
      PasswordInput.removeAttribute("disabled");
      PasswordInput.classList.remove("disabled");

      // > Change the text of the button
      Button.textContent = "Login";
    }
  }
}

async function SendingLoginInformation(Email, password) {

  // % Define the destination of transmission
  const destination = "/login";

  // % Listen for the response
  const response = await fetch(destination, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Cache-Control": "no-cache",
      Pragma: "no-cache",
    },
    body: JSON.stringify({ gmail: Email, password: password }),
  });

  if(!response.ok){
    throw new Error("Network response was not ok");
  }else{

    if(response.status == 200){
      window.location.href = "/excelConvert";
    }
  }

}
