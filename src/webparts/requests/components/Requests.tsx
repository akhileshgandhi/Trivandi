/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/explicit-function-return-type */

import * as React from "react";
import styles from "./Requests.module.scss";

import Hero from "../../../shared/component/Hero/Hero";
import type { IRequestsProps } from "./IRequestsProps";
import heroBg from "../../../shared/assets/hero_new.jpg";
import Section from "../../../shared/component/Section/Section";
import Card from "../../../shared/component/Card/Card";
import GlobalLoader from "../../../shared/component/GlobalLoader";
import "../../../shared/globalcss/globalcss.scss"
const requests = [
  {
    title: "Project Setup Request",
    subtitle: "Request a new CMAP project and folder structure.",
    color: "#ec2f91",
  },
  {
    title: "RFP Support Request",
    subtitle: "Ask for bid writing or graphic design support.",
    color: "#f58220",
  },
  {
    title: "Access Request",
    subtitle: "Request access to a restricted folder or site.",
    color: "#8fa3e8",
  },
  {
    title: "General Support",
    subtitle: "Ask a general question to the PMO team.",
    color: "#1327ff",
  },
];

const Requests: React.FC<IRequestsProps> = (props) => {
  const [isLoading, setIsLoading] = React.useState(true);
  
  React.useEffect(() => {
    setIsLoading(true);

    const timer = setTimeout(() => {
      setIsLoading(false);
    }, 600);

    return () => clearTimeout(timer);
  }, []);

  return (
    <div className={styles.page}>
      <Hero
        title={`Hey ${props.userDisplayName},`}
        subtitle="Step into your project control hub!"
        bg={heroBg}
      />
      {isLoading ? (
        <GlobalLoader variant="content" />
      ) : (
        <>
          {/* PAGE LABEL */}
          <div className={styles.label}>
            <span className={styles.labelLine}></span>
            <h4>Requests Hub</h4>
          </div>

          {/* CONTENT */}
          <div className={styles.outerBlock}>
            <Section>
              <div className={styles.requestsGrid}>
                {requests.map((req) => (
                  <Card
                    key={req.title}
                    title={req.title}
                    subtitle={req.subtitle}
                    color={req.color}
                    iconRight
                  />
                ))}
              </div>
            </Section>
          </div>
        </>
      )}
    </div>
  );
};

export default Requests;
