import * as React from 'react';
import styles from './Dashboard.module.scss';

interface EventCardProps {
  title: string;
  date: string;
  time: string;
  day: string;
  month: string;
}

const EventCard: React.FC<EventCardProps> = ({ title, date, time, day, month }) => {
  return (
    <div className={styles.eventCard}>
      <div className={styles.eventDate}>
        <div className={styles.eventMonth}>{month}</div>
        <div className={styles.eventDay}>{day}</div>
      </div>
      <div className={styles.eventDetails}>
        <h4 className={styles.eventTitle}>{title}</h4>
        <p className={styles.eventTime}>{date}, {time}</p>
      </div>
    </div>
  );
};

export default EventCard;
