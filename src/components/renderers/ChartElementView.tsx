import React from 'react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { ChartElement } from '../../types/pptx';

interface Props {
  element: ChartElement;
}

export const ChartElementView: React.FC<Props> = ({ element }) => {
  const { title, chartType, categories, series } = element;

  // Format data for Recharts
  const data = categories.map((cat, idx) => {
    const item: Record<string, any> = { category: cat };
    series.forEach((s) => {
      item[s.name] = s.values[idx] !== undefined ? s.values[idx] : 0;
    });
    return item;
  });

  return (
    <div className="w-full h-full bg-slate-900/90 text-white rounded-lg p-3 flex flex-col border border-slate-700/60 shadow-lg">
      {title && (
        <h4 className="text-sm font-semibold text-slate-200 mb-2 tracking-wide text-center">
          {title}
        </h4>
      )}

      <div className="flex-1 w-full min-h-0">
        <ResponsiveContainer width="100%" height="100%">
          {chartType === 'line' ? (
            <LineChart data={data} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
              <XAxis dataKey="category" stroke="#94A3B8" fontSize={11} />
              <YAxis stroke="#94A3B8" fontSize={11} />
              <Tooltip contentStyle={{ backgroundColor: '#1E293B', borderColor: '#475569' }} />
              <Legend wrapperStyle={{ fontSize: '11px', color: '#CBD5E1' }} />
              {series.map((s) => (
                <Line
                  key={s.name}
                  type="monotone"
                  dataKey={s.name}
                  stroke={s.color}
                  strokeWidth={2}
                  dot={{ r: 4 }}
                />
              ))}
            </LineChart>
          ) : chartType === 'pie' ? (
            <PieChart>
              <Tooltip contentStyle={{ backgroundColor: '#1E293B', borderColor: '#475569' }} />
              <Legend wrapperStyle={{ fontSize: '11px' }} />
              <Pie
                data={data}
                dataKey={series[0]?.name || 'category'}
                nameKey="category"
                cx="50%"
                cy="50%"
                outerRadius={65}
                label
              >
                {data.map((_, index) => (
                  <Cell
                    key={`cell-${index}`}
                    fill={series[index % series.length]?.color || '#3B82F6'}
                  />
                ))}
              </Pie>
            </PieChart>
          ) : (
            /* Bar / Column Chart */
            <BarChart data={data} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
              <XAxis dataKey="category" stroke="#94A3B8" fontSize={11} />
              <YAxis stroke="#94A3B8" fontSize={11} />
              <Tooltip contentStyle={{ backgroundColor: '#1E293B', borderColor: '#475569' }} />
              <Legend wrapperStyle={{ fontSize: '11px', color: '#CBD5E1' }} />
              {series.map((s) => (
                <Bar key={s.name} dataKey={s.name} fill={s.color} radius={[4, 4, 0, 0]} />
              ))}
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  );
};
