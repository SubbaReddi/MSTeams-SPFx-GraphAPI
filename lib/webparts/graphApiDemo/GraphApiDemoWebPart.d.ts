import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IGraphApiDemoWebPartProps {
    description: string;
}
export default class GraphApiDemoWebPart extends BaseClientSideWebPart<IGraphApiDemoWebPartProps> {
    private graphClient;
    onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=GraphApiDemoWebPart.d.ts.map