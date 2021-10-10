var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'GraphApiDemoWebPartStrings';
import GraphApiDemo from './components/GraphApiDemo';
var GraphApiDemoWebPart = /** @class */ (function (_super) {
    __extends(GraphApiDemoWebPart, _super);
    function GraphApiDemoWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    GraphApiDemoWebPart.prototype.onInit = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.context.msGraphClientFactory
                .getClient()
                .then(function (client) {
                _this.graphClient = client;
                resolve();
            }, function (err) { return reject(err); });
        });
    };
    GraphApiDemoWebPart.prototype.render = function () {
        var element = React.createElement(GraphApiDemo, {
            description: this.properties.description,
            graphClient: this.graphClient,
        });
        ReactDom.render(element, this.domElement);
    };
    GraphApiDemoWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(GraphApiDemoWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    GraphApiDemoWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return GraphApiDemoWebPart;
}(BaseClientSideWebPart));
export default GraphApiDemoWebPart;
//# sourceMappingURL=GraphApiDemoWebPart.js.map