import * as React from 'react';
export default class ErrorHandling extends React.Component {
    constructor(props) {
      super(props);
      this.state = { error: null, errorInfo: null };
    }
    
    componentDidCatch(error, errorInfo) {
      // Catch errors in any components below and re-render with error message
      this.setState({
        error: error,
        errorInfo: errorInfo
      })
      // You can also log error messages to an error reporting service here
    }
    
    render() {
      if (this.state.errorInfo) {
        // Error path
        return (
          <div>
            <h2>Something went wrong.</h2>
            <details style={{ whiteSpace: 'pre-wrap' }}>
              {this.state.error && this.state.error.toString()}
              <br />
              {this.state.errorInfo.componentStack}
            </details>
          </div>
        );
      }
      // Normally, just render children
      return this.props.children;
    }  
  }
  
// export default class ErrorHandling extends React.Component {
//     constructor(props) {
//         super(props);
//         this.state = {
//             hasError: false,
//             error: null
//         };
//     }
//     componentDidCatch(error, info) {
//         error ? console.log(error) : info ? console.log(info) : "";
//         this.setState({
//             hasError: true,
//             error: error
//          });
//     }
//     render() {
//         if (this.state.hasError) {
//             return <React.Fragment>
//                 <div className="error-block">
//                     <div className="inner">
//                         <Link to="/Listing">
//                             <i className="ms-Icon ms-Icon--Home home-btn"></i>
//                         </Link>
//                     </div>
//                     <h1>Oops!!! Something went wrong</h1>
//                 </div>
//                 {console.error(this.state.error.message)}
//             </React.Fragment>;
//         } else {
//             return this.props.children;
//         }
//     }
// }