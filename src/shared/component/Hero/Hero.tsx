import React from "react";
import styles from "./Hero.module.scss";

interface HeroProps {
  title: string;
  subtitle: string;
  bg: string;
}

const Hero: React.FC<HeroProps> = ({ title, subtitle, bg }) => {
  return (
    <div className={styles.hero} style={{ backgroundImage: `url(${bg})` }}>
      <div  className={styles.leftSpacing}>
      <h2>{title}</h2>
      <p>{subtitle}</p>
      </div>
    </div>
  );
};

export default Hero;
