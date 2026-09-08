async function checkJarvis() {

    const dot =
        document.getElementById("statusDot");

    const text =
        document.getElementById("statusText");


    try {

        const response =
            await fetch("/jarvis-status");

        const data =
            await response.json();


        if (data.connected) {

            dot.classList.add("online");

            text.textContent =
                "JARVIS online";

        } else {

            dot.classList.remove("online");

            text.textContent =
                "JARVIS offline";
        }

    } catch {

        dot.classList.remove("online");

        text.textContent =
            "JARVIS offline";
    }
}



function addMessage(role, message) {

    const messages =
        document.getElementById("messages");


    const emptyState =
        document.getElementById("emptyState");


    if (emptyState) {
        emptyState.remove();
    }


    const bubble =
        document.createElement("div");


    bubble.className =
        `msg ${role}`;


    const roleLabel =
        document.createElement("span");


    roleLabel.className =
        "role";


    roleLabel.textContent =
        role === "user"
            ? "You"
            : "JARVIS";


    const content =
        document.createElement("div");


    content.textContent =
        message;


    bubble.appendChild(
        roleLabel
    );


    bubble.appendChild(
        content
    );


    messages.appendChild(
        bubble
    );


    messages.scrollTop =
        messages.scrollHeight;
}



async function uploadPDF() {

    const fileInput =
        document.getElementById(
            "pdfFile"
        );


    const info =
        document.getElementById(
            "documentInfo"
        );


    const uploadButton =
        document.getElementById(
            "uploadBtn"
        );


    if (!fileInput.files.length) {

        info.textContent =
            "Choose a PDF first.";

        return;
    }


    const formData =
        new FormData();


    formData.append(
        "file",
        fileInput.files[0]
    );


    uploadButton.disabled =
        true;


    uploadButton.textContent =
        "Uploading...";


    info.textContent =
        "Reading your document...";


    try {

        const response =
            await fetch(
                "/upload",
                {
                    method: "POST",
                    body: formData
                }
            );


        const data =
            await response.json();


        if (data.error) {

            info.textContent =
                `Error: ${data.error}`;

            return;
        }


        info.textContent =
            `${data.filename}\n` +
            `${data.pages} pages · ` +
            `${data.characters_extracted.toLocaleString()} characters`;


        await clearChat(false);


    } catch (error) {

        info.textContent =
            `Upload failed: ${error.message}`;

    } finally {

        uploadButton.disabled =
            false;


        uploadButton.textContent =
            "Upload document";
    }
}



async function askJarvis() {

    const questionBox =
        document.getElementById(
            "question"
        );


    const askButton =
        document.getElementById(
            "askBtn"
        );


    const question =
        questionBox.value.trim();


    if (!question) {
        return;
    }


    addMessage(
        "user",
        question
    );


    questionBox.value =
        "";


    askButton.disabled =
        true;


    askButton.textContent =
        "Thinking...";


    const messages =
        document.getElementById(
            "messages"
        );


    const thinkingId =
        `thinking-${Date.now()}`;


    const thinking =
        document.createElement(
            "div"
        );


    thinking.id =
        thinkingId;


    thinking.className =
        "msg assistant";


    thinking.innerHTML = `
        <span class="role">
            JARVIS
        </span>

        <div>
            Reading the document...
        </div>
    `;


    messages.appendChild(
        thinking
    );


    messages.scrollTop =
        messages.scrollHeight;


    const formData =
        new FormData();


    formData.append(
        "question",
        question
    );


    try {

        const response =
            await fetch(
                "/ask",
                {
                    method: "POST",
                    body: formData
                }
            );


        const data =
            await response.json();


        document
            .getElementById(
                thinkingId
            )
            ?.remove();


        if (data.error) {

            addMessage(
                "assistant",
                `Error: ${data.error}`
            );

            return;
        }


        addMessage(
            "assistant",
            data.answer ||
            "JARVIS returned an empty response."
        );


    } catch (error) {

        document
            .getElementById(
                thinkingId
            )
            ?.remove();


        addMessage(
            "assistant",
            `Request failed: ${error.message}`
        );


    } finally {

        askButton.disabled =
            false;


        askButton.textContent =
            "Ask JARVIS";
    }
}



async function quickAsk(question) {

    document
        .getElementById(
            "question"
        )
        .value =
            question;


    await askJarvis();
}



async function clearChat(
    clearQuestion = true
) {

    try {

        await fetch(
            "/clear-chat",
            {
                method: "POST"
            }
        );

    } catch {
        // The visible chat can still
        // be cleared if this fails.
    }


    const messages =
        document.getElementById(
            "messages"
        );


    messages.innerHTML = `
        <div
            id="emptyState"
            class="empty-state"
        >

            <div class="empty-mark">
                J
            </div>

            <h3>
                Ready when you are.
            </h3>

            <p>
                Upload a PDF, then ask
                JARVIS anything about it.
            </p>

        </div>
    `;


    if (clearQuestion) {

        document
            .getElementById(
                "question"
            )
            .value =
                "";
    }
}



document
    .getElementById(
        "question"
    )
    .addEventListener(
        "keydown",
        function (event) {

            if (
                event.key === "Enter" &&
                !event.shiftKey
            ) {

                event.preventDefault();

                askJarvis();
            }
        }
    );



checkJarvis();


setInterval(
    checkJarvis,
    10000
);