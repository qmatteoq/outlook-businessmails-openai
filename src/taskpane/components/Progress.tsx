import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

export interface ProgressProps {
  logo?: string;
  message: string;
  title: string;
}

export default class Progress extends React.Component<ProgressProps> {
  render() {
    const { logo, message, title } = this.props;

    const logoSection = () => {
      if (logo === undefined) {
        return <div></div>;
      } else {
        return <img width="90" height="90" src={logo} alt={title} title={title} />;
      }
    };

    const titleSection = () => {
      if (title === undefined) {
        return <div></div>;
      } else {
        return <h1 className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{title}</h1>;
      }
    };

    return (
      <section className="ms-welcome__progress ms-u-fadeIn500">
        {logoSection}
        {titleSection}
        <Spinner size={SpinnerSize.large} label={message} />
      </section>
    );
  }
}
