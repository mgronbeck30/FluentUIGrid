import * as React from 'react';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DetailsList, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { CommandBarButton, IIconProps, Stack, IStackStyles } from 'office-ui-fabric-react';

export interface IDataSetProps {
    cols: IColumn[],
    data: object[],
    onButtonClicked? : () => void;
}

const completeIcon: IIconProps = { iconName: 'Contact' };
const stackStyles: Partial<IStackStyles> = { root: { height: 50, paddingLeft: 20 } };
export const HoverCardBasicExample: React.FunctionComponent<IDataSetProps> = props => {
    const {cols,data,onButtonClicked} = props;
    return(
  <Fabric>
      <Stack horizontal styles={stackStyles}>
      <CommandBarButton
        iconProps={completeIcon}
        text="WhoAmI"
        disabled={false}
        checked={false}
        onClick={onButtonClicked} 
      /></Stack>
    <DetailsList setKey="hoverSet" items={data} columns={cols}/>
  </Fabric>
    );
};