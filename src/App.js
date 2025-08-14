import React from "react";
import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import { Provider } from "react-redux";
import store from "./store/store";

import FinalUIDAIDashboard from "./pages/Dashboard";
import EnhancedAmcCalculator from "./pages/EnhancedAmcCalculator";
import AmcPaymentTracker from "./pages/AmcPaymentTracker";
import WarrantyEstimator from "./pages/WarrantyEstimator";
import WarrantyPaymentTracker from "./pages/WarrantyPaymentTracker";

import "./index.css";

function App() {
  return (
    <Provider store={store}>
      <Router>
        <div className="App">
          <Routes>
            <Route path="/" element={<FinalUIDAIDashboard />} />
            <Route
              path="/enhanced-amc-calculator"
              element={<EnhancedAmcCalculator />}
            />
            
            <Route
              path="/warranty-estimator"
              element={<WarrantyEstimator />}
            />

            <Route path="/amc-payment-tracker" element={<AmcPaymentTracker />} />
            <Route path="/warranty-payment-tracker" element={<WarrantyPaymentTracker />} />
          </Routes>
        </div>
      </Router>
    </Provider>
  );
}

export default App;
