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
import styles from './GraphApiDemo.module.scss';
var GraphApiDemo = /** @class */ (function (_super) {
    __extends(GraphApiDemo, _super);
    function GraphApiDemo(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            messages: [{
                    subject: ''
                }]
        };
        return _this;
    }
    GraphApiDemo.prototype.getDriveItems = function () {
        var _this = this;
        var getMessages = "me/messages";
        if (!this.props.graphClient) {
            return;
        }
        this.props.graphClient
            .api(getMessages)
            .version("v1.0")
            .select("subject,sentDateTime,webLink")
            .top(5)
            .get(function (err, res) {
            if (err) {
                console.log("Getting error in retrieving mesages =>", err);
            }
            if (res) {
                console.log("Success");
                if (res && res.value.length) {
                    console.log(res.value);
                    _this.setState({
                        messages: res.value
                    });
                }
            }
        });
    };
    GraphApiDemo.prototype.componentDidMount = function () {
        this.getDriveItems();
    };
    GraphApiDemo.prototype.render = function () {
        return (React.createElement("div", { className: styles.graphApiDemo }, this.state.messages.map(function (m) { return React.createElement(React.Fragment, null,
            React.createElement("span", null, m.subject),
            React.createElement("br", null)); })));
    };
    return GraphApiDemo;
}(React.Component));
export default GraphApiDemo;
//# sourceMappingURL=GraphApiDemo.js.map