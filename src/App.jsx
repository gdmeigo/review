// gdm-review.jsx を同じ src/ フォルダに置いてください。
// (中でも default export の関数名が "App" になっているので、
//  ここでは名前を変えてインポートしています)
import GdmReview from "./gdm-review.jsx";
import "./App.css";

function App() {
  return <GdmReview />;
}

export default App;
