
import * as React from 'react';
import { useEffect, useState } from 'react';
import type { IDashboardProps } from './IDashboardProps';
import styles from './Dashboard.module.scss';
import '../../../shared/globalcss/globalcss.scss';
import {
    getProjectStats,
    getComplianceItems,
    getKeyPeople,
    getUpcomingEvents,
    getKeyTools,
    getQuickLinks,
    getChildLinksByParent
} from '../../../shared/services/dashboardService';
// Import sub-components
import StatCard from './StatCard';
import ComplianceCard from './ComplianceCard';
import KeyPersonCard from './KeyPersonCard';
import EventCard from './EventCard';
import ToolCard from './ToolCard';
import QuickLinkCard from './QuickLinkCard';
import GlobalLoader from '../../../shared/component/GlobalLoader';
import MyWorkNew from './MyWorkNew';
import Hero from '../../../shared/component/Hero/Hero';
import heroBg from '../../../shared/assets/hero_new.jpg';
interface ProjectStats {
    totalProjects: number;
    liveProjects: number;
    myProjects: number;
    myBids: number;
}

interface ComplianceItem {
    id: number;
    title: string;
    description: string;
    color: string;
    icon: string;
    link?: string;
}

interface KeyPerson {
    id: number;
    name: string;
    title: string;
    imageUrl: string;
    email: string;
    phone: string;
}

interface Event {
    id: number;
    title: string;
    date: string;
    time: string;
    day: string;
    month: string;
}

interface Tool {
    id: number;
    title: string;
    description: string;
    color: string;
    icon: string;
    link: string;
}

interface QuickLink {
    id: number;
    title: string;
    url: string;
    backgroundColor: string;
    icon: string;
    hasChildren?: boolean;
}

const Dashboard: React.FC<IDashboardProps> = (props) => {
    const [isLoading, setIsLoading] = useState(true);
    const [projectStats, setProjectStats] = useState<ProjectStats>({
        totalProjects: 0,
        liveProjects: 0,
        myProjects: 0,
        myBids: 0
    });
    const [complianceItems, setComplianceItems] = useState<ComplianceItem[]>([]);
    const [keyPeople, setKeyPeople] = useState<KeyPerson[]>([]);
    const [upcomingEvents, setUpcomingEvents] = useState<Event[]>([]);
    const [keyTools, setKeyTools] = useState<Tool[]>([]);
    const [quickLinks, setQuickLinks] = useState<QuickLink[]>([]);

    const firstCardicon = require('../../../shared/assets/folderopen.png');
    const secondCardicon = require('../../../shared/assets/flash.png');
    const thirdCardicon = require('../../../shared/assets/direct.png');
    const fourthCardicon = require('../../../shared/assets/books.png');


    const loadDashboardData = async (): Promise<void> => {
        try {
            setIsLoading(true);

            // Load all data in parallel
            const [stats, compliance, people, events, links] = await Promise.all([
                getProjectStats(),
                getComplianceItems(),
                getKeyPeople(),
                getUpcomingEvents(),
                // getKeyTools(),
                getQuickLinks()
            ]);

            setProjectStats(stats);
            setComplianceItems(compliance);
            setKeyPeople(people);
            setUpcomingEvents(events);
            // setKeyTools(tools);
            setQuickLinks(links);
        } catch (error) {
            console.error('Error loading dashboard data:', error);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        // eslint-disable-next-line no-void
        void loadDashboardData();
    }, []);

    if (isLoading) {
        return <GlobalLoader />;
    }



    return (
        <div className={styles.dashboard}>
            {/* Hero Banner */}
            {/* <div className={styles.heroBanner} style={{ position: 'relative', overflow: 'hidden' }}>
                <video
                    autoPlay
                    muted
                    loop
                    playsInline
                    style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', objectFit: 'cover' }}
                >
                    <source src={vdo} type="video/mp4" />
                </video>
                <h1 className={styles.heroTitle} style={{ position: 'relative', zIndex: 1 }}>Hey {props.userDisplayName},</h1>
                <p className={styles.heroSubtitle} style={{ position: 'relative', zIndex: 1 }}>Step into your project control hub!</p>
            </div> */}
            <div className={styles.heroBanner}>
                <Hero
                    title={`Hey ${props.userDisplayName}`}
                    subtitle="Step into your project control hub!"
                    bg={heroBg}
                />
            </div>


            <div className={styles.mainContentQuickLinks}>

                {/* Quick Links Section */}
              <div className={styles.quickLinksSection}>
                        {quickLinks.length === 0 ? (
                            <div className={styles.emptyState}>
                                <img src={require('../../../shared/assets/Empty.png')} alt="No data" className={styles.emptyIcon} />
                                <p className={styles.emptyText}>No quick links available</p>
                            </div>
                        ) : (
                            <div className={styles.quickLinksGrid}>
                                {quickLinks.map((link) => (
                                    <QuickLinkCard
                                        key={link.id}
                                        title={link.title}
                                        url={link.url}
                                        backgroundColor={link.backgroundColor}
                                        icon={link.icon}
                                        hasChildren={link.hasChildren}
                                        fetchChildren={getChildLinksByParent}
                                    />
                                ))}
                            </div>
                        )}
                    </div>

                {/* <div className={styles.rightColumn}>
                    <div className={styles.sectionHeader}>
                        <span className={styles.accentBar}></span>
                        <h2 className={styles.sectionTitle}>Key Tools</h2>
                    </div> 
                     Key Tools Section*
                    <div className={styles.quickLinksSection}>
                        {keyTools.length === 0 ? (
                            <div className={styles.emptyState}>
                                <img src={require('../../../shared/assets/Empty.png')} alt="No data" className={styles.emptyIcon} />
                                <p className={styles.emptyText}>No key tools available</p>
                            </div>
                        ) : (
                            <div className={styles.toolsGrid}>
                                {keyTools.map((tool) => (
                                    <ToolCard
                                        key={tool.id}
                                        title={tool.title}
                                        description={tool.description}
                                        color={tool.color}
                                        icon={tool.icon}
                                        link={tool.link}
                                    />
                                ))}
                            </div>
                        )}
                    </div>
                </div> */}
            </div>
             {/* Project Stats Section */}
            <div className={styles.statsSection}>
                {/* <StatCard
                    title="Total Projects"
                    value={projectStats.totalProjects}
                    icon={firstCardicon}
                    color="#0511F2"
                /> */}
                <StatCard
                    title="My Projects"
                    value={projectStats.myProjects}
                    icon={thirdCardicon}
                    color="#0511F2"
                />
                <StatCard
                    title="Live Projects"
                    value={projectStats.liveProjects}
                    icon={secondCardicon}
                    color="#0511F2"
                />

                <StatCard
                    title="Active Pipeline"
                    value={projectStats.myBids}
                    icon={fourthCardicon}
                    color="#0511F2"
                />
            </div>
            {/* Standards & Compliance Section */}
            {/* <div className={styles.sectionHeader}>
                <span className={styles.accentBar}></span>
                <h2 className={styles.sectionTitle}>Standards &amp; Compliance</h2>
            </div>
            <div className={styles.section}>
                {complianceItems.length === 0 ? (
                    <div className={styles.emptyState}>
                        <img src={require('../../../shared/assets/Empty.png')} alt="No data" className={styles.emptyIcon} />
                        <p className={styles.emptyText}>No compliance items available</p>
                    </div>
                ) : (
                    <div className={styles.complianceGrid}>
                        {complianceItems.map((item) => (
                            <ComplianceCard
                                key={item.id}
                                title={item.title}
                                description={item.description}
                                color={item.color}
                                icon={item.icon}
                                link={item.link}
                            />
                        ))}
                    </div>
                )}
            </div> */}

            {/* Two Column Layout - Key People & Coming Up */}


            <div className={styles.mainContent} style={{ display: 'block' }}>
                <MyWorkNew props={props} />
            </div>

            {/* <div className={styles.mainContent} style={{ display: 'block' }}>
                <div className={styles.sectionHeader}>
                    <span className={styles.accentBar}></span>
                    <h2 className={styles.sectionTitle}>Key Tools</h2>
                </div>

                <div className={styles.section}>
                    {keyTools.length === 0 ? (
                        <div className={styles.emptyState}>
                            <img src={require('../../../shared/assets/Empty.png')} alt="No data" className={styles.emptyIcon} />
                            <p className={styles.emptyText}>No key tools available</p>
                        </div>
                    ) : (
                        <div className={styles.toolsGrid}>
                            {keyTools.map((tool) => (
                                <ToolCard
                                    key={tool.id}
                                    title={tool.title}
                                    description={tool.description}
                                    color={tool.color}
                                    icon={tool.icon}
                                />
                            ))}
                        </div>
                    )}
                </div>
            </div> */}
        </div>
    );
};

export default Dashboard;