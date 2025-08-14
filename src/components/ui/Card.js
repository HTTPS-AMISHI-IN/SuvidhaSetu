import React from "react";
import { Link } from "react-router-dom";

export function Card({ title, desc, to }) {
  return (
    <div className="bg-white text-gray-900 p-6 rounded-2xl shadow-xl transition-transform hover:-translate-y-1 hover:shadow-2xl">
      <h2 className="text-xl font-semibold text-blue-800 mb-2">{title}</h2>
      <p className="text-gray-700 text-sm">{desc}</p>
      <Link to={to} className="inline-block mt-4 px-4 py-2 bg-blue-900 text-white rounded-md hover:bg-slate-900 text-sm">
        Launch Tool
      </Link>
    </div>
  );
}

export function CardContent({ children }) {
  return <div>{children}</div>;
}
