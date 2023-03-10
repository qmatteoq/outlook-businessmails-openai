import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
// import { Configuration, OpenAIApi } from "openai";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
  isLoading: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);
    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
      isLoading: false,
    };
  }

  generateText = async () => {
    // eslint-disable-next-line no-undef
    var current = this;

    const apiKey = "5X5l_tzKawUtPkD7MuK-Mw-qW0r6aQdR6okgQWi0NOnLAzFuGzrtbA==";
    const endpoint = "https://openaibusinessgeneratorfunction.azurewebsites.net/";
    const userPrompt = this.state.startText;

    const url = endpoint + "api/AzureOpenAiEndpoint?code=" + apiKey;

    current.setState({ isLoading: true });

    const payload = {
      prompt: userPrompt,
    };

    // eslint-disable-next-line no-undef
    var response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    var data = await response.json();
    current.setState({ isLoading: false });
    current.setState({ generatedText: data.text });
  };

  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
    const isLoading = this.state.isLoading;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    const ProgressSection = () => {
      if (isLoading) {
        return <Progress title="Loading..." message="Generating the message." />;
      } else {
        return <div> </div>;
      }
    };

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Open AI business e-mail generator
          </h2>
          <p>Briefly describe what you want to communicate in the mail:</p>
          <textarea
            className="ms-welcome"
            onChange={(e) => this.setState({ startText: e.target.value })}
            rows={5}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.generateText}
            >
              Generate text
            </DefaultButton>
          </p>
          <p>
            <ProgressSection />
          </p>
          <textarea
            className="ms-welcome"
            defaultValue={this.state.generatedText}
            onChange={(e) => this.setState({ finalMailText: e.target.value })}
            rows={15}
            cols={40}
          />
          <p>
            <DefaultButton
              className="ms-welcome__action"
              iconProps={{ iconName: "ChevronRight" }}
              onClick={this.insertIntoMail}
            >
              Insert into mail
            </DefaultButton>
          </p>
        </main>
      </div>
    );
  }
}
