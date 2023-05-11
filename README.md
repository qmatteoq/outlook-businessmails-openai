# Outlook OpenAI Add-in

This add-in allows you to write professional business emails with the help of OpenAI APIs. You can simply type a simple sentence and the add-in will generate a polished and formal email based on your input.

The project comes into two variants:

- the master branch is based on the Open AI APIs, so you will need to get an API Key from the [OpenAI portal](https://platform.openai.com/overview).
- the azure-open-ai branch is based on the [Azure Open AI APIs](https://azure.microsoft.com/en-us/products/cognitive-services/openai-service), so you will need to create a service instance on your Azure subscription. Please be aware that you must be approved by Microsoft before being able to use Azure Open AI.

**Please note:** this project is meant to help developers understanding how they can integrate OpenAI APIs and ChatGPT into their applications, in this case an Outlook add-in. It's not meant to be a "ready to be deployed" solution.

This project is a companion of the following articles, which explains how this add-in was built:

- [Bringing OpenAI into an Outlook add-in: a business mail generator](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/bringing-openai-into-an-outlook-add-in-a-business-mail-generator/ba-p/3743099)
- [Bringing Open AI into an Outlook add-in: moving to Azure Open AI](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/bringing-openai-into-an-outlook-add-in-a-business-mail-generator/ba-p/3743099)
- [Bring the ChatGPT model into our applications](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/bring-the-chatgpt-model-into-our-applications/ba-p/3766574)

## How to use

**Please note**: make sure to read [the companion blog post](https://techcommunity.microsoft.com/t5/modern-work-app-consult-blog/bringing-openai-into-an-outlook-add-in-a-business-mail-generator/ba-p/3743099) to understand all the requirements.

- Clone the repository on your machine
- Open the solution with Visual Studio Code.
- Make sure to replace the existing placeholders with your information:
  - In case you're using the master branch, add your API key
  - In case you're using Azure Open AI service, add your URL, deployment name and API key.
- Move to the Debug tab in Visual Studio Code, choose **Outlook Desktop (Edge Chromium)** and press F5. The add-in will be sideloaded in Outlook desktop.
- Compose a new mail, click on AI Assistant, type one or two sentences and click on the Generate button and wait for the add-in to produce a professional email. You can edit the email as you wish before sending it.
- Enjoy the convenience and efficiency of writing business emails with OpenAI!

## Features

- The add-in uses OpenAI's GPT-3 language model to generate high-quality and natural-sounding emails.
- The add-in adapts to different contexts and tones based on your input and the recipient's information.
- The add-in respects your privacy and does not store or share your email content.

## Deployment

This add-in is built using the Office web model. If you want to deploy it on other machines without using Visual Studio Code, you must host it on a web storage and change the manifest file to point to the new URL. Refer to [the following documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code) to learn how to host the add-in on Azure Storage.

## Feedback

If you have any feedback, suggestions or issues with the add-in, please feel free to open an issue or a pull request on this repository.
