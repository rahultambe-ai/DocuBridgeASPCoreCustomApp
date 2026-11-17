import * as React from 'react';
import styles from './Dms.module.scss';
import AnalyticsDashboard from './UserPortal/AnalyticsDashboard';

export interface IDmsProps {
  context: any;
}

export default function Dms({ context }: IDmsProps): React.ReactElement<IDmsProps> {
  return (
    <section className={styles.dms}>
      <AnalyticsDashboard context={context} />
    </section>
  );
}
