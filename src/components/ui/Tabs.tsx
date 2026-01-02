import type { ReactNode } from 'react';
import { Link } from 'react-router-dom';

interface Tab {
  id: string;
  label: string;
  icon?: ReactNode;
}

interface TabsProps {
  tabs: Tab[];
  activeTab: string;
  onChange?: (tabId: string) => void;
}

export function Tabs({ tabs, activeTab }: TabsProps) {
  return (
    <div className="border-b border-gray-200 dark:border-slate-700">
      <nav className="flex gap-1" aria-label="Tabs">
        {tabs.map((tab) => (
          <Link
            key={tab.id}
            to={`/${tab.id}`}
            className={`
              flex items-center gap-2 px-4 py-3 text-sm font-medium
              border-b-2 transition-colors duration-200
              ${activeTab === tab.id
                ? 'border-indigo-500 text-indigo-600 dark:text-indigo-400'
                : 'border-transparent text-gray-500 dark:text-slate-400 hover:text-gray-700 dark:hover:text-slate-200 hover:border-gray-300 dark:hover:border-slate-500'
              }
            `}
          >
            {tab.icon}
            {tab.label}
          </Link>
        ))}
      </nav>
    </div>
  );
}
