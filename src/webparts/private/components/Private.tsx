import * as React from 'react';
import { IPrivateProps } from './IPrivateProps';
import PrivateComponent from '../CustomComponents/PrivateComponent';
import '../../../shared/globalcss/globalcss.scss';
const Private: React.FC<IPrivateProps> = (props) => {
  return (
    <div className="CanvasZoneSectionContainer">
      <PrivateComponent context={props.context} />
    </div>

  );
};
export default Private;