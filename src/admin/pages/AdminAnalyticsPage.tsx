import { useState, useEffect, useMemo } from 'react';
import { firebaseGetActivity, firebaseGetUsers } from '../../lib/firebaseAuth';
import { BarChart3, User, Calendar, RefreshCw, ChevronDown, TrendingUp, Clock, Edit3, Download, Shield, LogIn, X } from 'lucide-react';

interface UserStat {
  id: string; name: string; email: string;
  totalActions: number; logins: number; edits: number;
  downloads: number; permChanges: number;
  lastActive: string;
  actionsByDay: Record<string, number>;
}

const ACTION_COLORS: Record<string,string> = {
  login:'#10b981', admin_login:'#ef4444', edit_cells:'#6366f1',
  add_row:'#3b82f6', delete_row:'#f43f5e', add_column:'#8b5cf6',
  delete_column:'#f59e0b', download_data:'#06b6d4', update_permissions:'#ec4899',
  bulk_delete_rows:'#ef4444', logout:'#94a3b8', other:'#a78bfa'
};
const ACTION_LABELS: Record<string,string> = {
  login:'Login', admin_login:'Admin Login', edit_cells:'Cell Edits',
  add_row:'Add Row', delete_row:'Delete Row', add_column:'Add Column',
  delete_column:'Del Column', download_data:'Download', update_permissions:'Permissions',
  bulk_delete_rows:'Bulk Delete', logout:'Logout', update_user:'Update User',
  change_password:'Change Pass', create_user:'Create User', delete_user:'Delete User',
  download_request:'DL Request', respond_download_request:'DL Response'
};

// Pure CSS bar chart component
function BarChart({ data, label }: { data: {key:string;value:number;color:string}[]; label:string }) {
  const max = Math.max(...data.map(d=>d.value), 1);
  if (data.length === 0) return <div style={{padding:'30px',textAlign:'center',color:'var(--muted)',fontSize:'13px'}}>No data</div>;
  return (
    <div>
      <div style={{fontSize:'12px',fontWeight:700,color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.04em',marginBottom:'12px'}}>{label}</div>
      <div style={{display:'flex',alignItems:'flex-end',gap:'3px',height:'140px',padding:'0 4px'}}>
        {data.map((d,i)=>(
          <div key={i} style={{flex:1,display:'flex',flexDirection:'column',alignItems:'center',gap:'4px',minWidth:0}}>
            <span style={{fontSize:'10px',fontWeight:700,color:'var(--foreground)'}}>{d.value||''}</span>
            <div style={{width:'100%',maxWidth:'32px',background:d.color,borderRadius:'4px 4px 0 0',
              height:`${Math.max((d.value/max)*110,2)}px`,transition:'height 0.5s cubic-bezier(.4,0,.2,1)',opacity:0.85}} />
            <span style={{fontSize:'9px',color:'var(--muted)',fontWeight:600,whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis',maxWidth:'40px',textAlign:'center'}}>{d.key}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// Donut chart
function DonutChart({ data }: { data: {label:string;value:number;color:string}[] }) {
  const total = data.reduce((s,d)=>s+d.value,0) || 1;
  let cum = 0;
  const segments = data.filter(d=>d.value>0).map(d=>{
    const pct = (d.value/total)*100;
    const start = cum; cum += pct;
    return {...d, pct, start};
  });
  const gradient = segments.map(s=>`${s.color} ${s.start}% ${s.start+s.pct}%`).join(', ');
  return (
    <div style={{display:'flex',alignItems:'center',gap:'20px',flexWrap:'wrap'}}>
      <div style={{width:'110px',height:'110px',borderRadius:'50%',background:`conic-gradient(${gradient||'var(--border) 0% 100%'})`,
        display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
        <div style={{width:'60px',height:'60px',borderRadius:'50%',background:'var(--surface)',display:'flex',alignItems:'center',justifyContent:'center',
          fontSize:'16px',fontWeight:800,color:'var(--foreground)'}}>{data.reduce((s,d)=>s+d.value,0)}</div>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:'4px',flex:1,minWidth:'120px'}}>
        {segments.slice(0,8).map((s,i)=>(
          <div key={i} style={{display:'flex',alignItems:'center',gap:'6px',fontSize:'12px'}}>
            <div style={{width:'10px',height:'10px',borderRadius:'3px',background:s.color,flexShrink:0}} />
            <span style={{color:'var(--muted)',fontWeight:500,flex:1}}>{s.label}</span>
            <span style={{fontWeight:700,color:'var(--foreground)'}}>{s.value}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

const sty = {
  card: {background:'var(--surface)',borderRadius:'12px',border:'1px solid var(--border)',padding:'16px 20px',boxShadow:'0 1px 3px rgba(0,0,0,.06)'} as React.CSSProperties,
  statCard: {background:'var(--surface)',borderRadius:'10px',border:'1px solid var(--border)',padding:'14px 16px',boxShadow:'0 1px 2px rgba(0,0,0,.04)',flex:'1 1 120px',minWidth:'110px'} as React.CSSProperties,
  label: {fontSize:'11px',fontWeight:700,color:'var(--muted)',textTransform:'uppercase' as const,letterSpacing:'0.04em',marginBottom:'4px',display:'block'},
  select: {width:'100%',padding:'8px 28px 8px 10px',borderRadius:'6px',border:'1px solid var(--border)',background:'var(--background)',color:'var(--foreground)',fontSize:'13px',fontWeight:500,appearance:'none' as const,cursor:'pointer',outline:'none'},
  dateInput: {width:'100%',padding:'7px 10px',borderRadius:'6px',border:'1px solid var(--border)',background:'var(--background)',color:'var(--foreground)',fontSize:'13px',outline:'none',boxSizing:'border-box' as const},
};

export default function AdminAnalyticsPage() {
  const [activities, setActivities] = useState<any[]>([]);
  const [users, setUsers] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [filterUser, setFilterUser] = useState<string>('all');
  const [filterInterval, setFilterInterval] = useState<string>('7d');
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [filterSingleDate, setFilterSingleDate] = useState('');

  const fetch_ = async () => {
    setLoading(true);
    try {
      const [a, u] = await Promise.all([firebaseGetActivity(2000), firebaseGetUsers()]);
      setActivities(a.activities || []);
      setUsers((u.users || []).filter((x: any) => x.role !== 'superadmin'));
    } catch {} finally { setLoading(false); }
  };
  useEffect(() => { fetch_(); }, []);

  const dateRange = useMemo(() => {
    if (filterSingleDate) {
      const d = new Date(filterSingleDate);
      const from = new Date(d); from.setHours(0,0,0,0);
      const to = new Date(d); to.setHours(23,59,59,999);
      return { from, to };
    }
    if (filterDateFrom && filterDateTo) return { from: new Date(filterDateFrom), to: new Date(filterDateTo + 'T23:59:59') };
    const to = new Date(); to.setHours(23,59,59,999);
    const from = new Date();
    if (filterInterval === '1d') from.setDate(from.getDate() - 1);
    else if (filterInterval === '7d') from.setDate(from.getDate() - 7);
    else if (filterInterval === '30d') from.setDate(from.getDate() - 30);
    else if (filterInterval === '90d') from.setDate(from.getDate() - 90);
    else from.setFullYear(from.getFullYear() - 10);
    from.setHours(0,0,0,0);
    return { from, to };
  }, [filterInterval, filterDateFrom, filterDateTo, filterSingleDate]);

  const rangedActivities = useMemo(() => {
    return activities.filter(a => {
      const d = new Date(a.timestamp);
      if (d < dateRange.from || d > dateRange.to) return false;
      if (filterUser !== 'all' && a.userId !== filterUser) return false;
      return true;
    });
  }, [activities, dateRange, filterUser]);

  // Per-user stats
  const userStats = useMemo<UserStat[]>(() => {
    const map = new Map<string, UserStat>();
    users.forEach(u => map.set(u.id, { id:u.id, name:u.name, email:u.email, totalActions:0, logins:0, edits:0, downloads:0, permChanges:0, lastActive:'', actionsByDay:{} }));
    rangedActivities.forEach(a => {
      if (!map.has(a.userId)) map.set(a.userId, { id:a.userId, name:a.userName||'Unknown', email:'', totalActions:0, logins:0, edits:0, downloads:0, permChanges:0, lastActive:'', actionsByDay:{} });
      const s = map.get(a.userId)!;
      s.totalActions++;
      if (a.action==='login'||a.action==='admin_login') s.logins++;
      if (['edit_cells','add_row','delete_row','bulk_delete_rows','add_column','delete_column','update_user','update_permissions'].includes(a.action)) s.edits++;
      if (['download_data','download_request','respond_download_request'].includes(a.action)) s.downloads++;
      if (a.action==='update_permissions') s.permChanges++;
      if (!s.lastActive||a.timestamp>s.lastActive) s.lastActive = a.timestamp;
      const day = new Date(a.timestamp).toISOString().split('T')[0];
      s.actionsByDay[day] = (s.actionsByDay[day]||0)+1;
    });
    return Array.from(map.values()).sort((a,b)=>b.totalActions-a.totalActions);
  }, [users, rangedActivities]);

  // Daily activity data for bar chart
  const dailyData = useMemo(() => {
    const map = new Map<string,number>();
    rangedActivities.forEach(a => {
      const day = new Date(a.timestamp).toLocaleDateString('en-US',{month:'short',day:'numeric'});
      map.set(day, (map.get(day)||0)+1);
    });
    const entries = Array.from(map.entries());
    const last14 = entries.slice(-14);
    return last14.map(([key,value])=>({key,value,color:'var(--navy)'}));
  }, [rangedActivities]);

  // Action type distribution for donut
  const actionDist = useMemo(() => {
    const map = new Map<string,number>();
    rangedActivities.forEach(a => map.set(a.action, (map.get(a.action)||0)+1));
    return Array.from(map.entries())
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>({label:ACTION_LABELS[k]||k.replace(/_/g,' '),value:v,color:ACTION_COLORS[k]||ACTION_COLORS.other}));
  }, [rangedActivities]);

  // Hourly distribution
  const hourlyData = useMemo(() => {
    const hours = new Array(24).fill(0);
    rangedActivities.forEach(a => { hours[new Date(a.timestamp).getHours()]++; });
    return hours.map((v,i)=>({key:`${i}`,value:v,color:i>=9&&i<=18?'#10b981':'#6366f1'}));
  }, [rangedActivities]);

  // Per-user bar data
  const userBarData = useMemo(() => {
    return userStats.filter(u=>u.totalActions>0).slice(0,10).map((u,i)=>({
      key:u.name.split(' ')[0],value:u.totalActions,
      color:`hsl(${(i*45)%360},65%,55%)`
    }));
  }, [userStats]);

  const displayUsers = filterUser === 'all' ? userStats : userStats.filter(u => u.id === filterUser);
  const topWorker = userStats.length > 0 ? userStats[0] : null;
  const activeFilterCount = [filterUser!=='all', filterInterval!=='7d'||!!filterSingleDate, !!filterDateFrom, !!filterDateTo, !!filterSingleDate].filter(Boolean).length;

  const clearFilters = () => { setFilterUser('all'); setFilterInterval('7d'); setFilterDateFrom(''); setFilterDateTo(''); setFilterSingleDate(''); };

  return (
    <div>
      {/* Header */}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'16px',flexWrap:'wrap',gap:'10px'}}>
        <h2 style={{margin:0,fontSize:'18px',fontWeight:700,color:'var(--foreground)',display:'flex',alignItems:'center',gap:'10px'}}>
          <BarChart3 size={20} color="var(--navy)" /> Employee Analytics
        </h2>
        <div style={{display:'flex',gap:'8px'}}>
          {activeFilterCount > 0 && <button onClick={clearFilters} style={{background:'var(--destructive-bg)',border:'1px solid var(--border)',color:'var(--destructive)',cursor:'pointer',padding:'8px 14px',borderRadius:'8px',display:'flex',alignItems:'center',gap:'4px',fontSize:'12px',fontWeight:600}}><X size={14}/>Clear</button>}
          <button onClick={fetch_} style={{background:'var(--surface)',border:'1px solid var(--border)',color:'var(--navy)',cursor:'pointer',padding:'10px',borderRadius:'8px',display:'flex',boxShadow:'0 1px 2px rgba(0,0,0,.04)'}}><RefreshCw size={16}/></button>
        </div>
      </div>

      {/* Filters */}
      <div style={{display:'flex',flexWrap:'wrap',gap:'12px',marginBottom:'16px',...sty.card,padding:'12px 16px'}}>
        <div style={{flex:'1 1 150px',minWidth:'130px'}}>
          <label style={sty.label}><User size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> Employee</label>
          <div style={{position:'relative'}}>
            <select value={filterUser} onChange={e=>setFilterUser(e.target.value)} style={sty.select}>
              <option value="all">All Employees</option>
              {users.map((u:any)=><option key={u.id} value={u.id}>{u.name}</option>)}
            </select>
            <ChevronDown size={12} style={{position:'absolute',right:'8px',top:'50%',transform:'translateY(-50%)',pointerEvents:'none',color:'var(--muted)'}}/>
          </div>
        </div>
        <div style={{flex:'1 1 130px',minWidth:'110px'}}>
          <label style={sty.label}><Clock size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> Interval</label>
          <div style={{position:'relative'}}>
            <select value={filterInterval} onChange={e=>{setFilterInterval(e.target.value);setFilterDateFrom('');setFilterDateTo('');setFilterSingleDate('');}} style={sty.select}>
              <option value="1d">Today</option>
              <option value="7d">Last 7 Days</option>
              <option value="30d">Last 30 Days</option>
              <option value="90d">Last 90 Days</option>
              <option value="all">All Time</option>
            </select>
            <ChevronDown size={12} style={{position:'absolute',right:'8px',top:'50%',transform:'translateY(-50%)',pointerEvents:'none',color:'var(--muted)'}}/>
          </div>
        </div>
        <div style={{flex:'1 1 130px',minWidth:'120px'}}>
          <label style={sty.label}><Calendar size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> Specific Date</label>
          <input type="date" value={filterSingleDate} onChange={e=>{setFilterSingleDate(e.target.value);if(e.target.value){setFilterDateFrom('');setFilterDateTo('');setFilterInterval('custom');}}} style={{...sty.dateInput,border:filterSingleDate?'1.5px solid var(--navy)':'1px solid var(--border)',background:filterSingleDate?'rgba(30,41,82,0.04)':'var(--background)'}} />
        </div>
        <div style={{flex:'1 1 110px',minWidth:'100px'}}>
          <label style={sty.label}><Calendar size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> From</label>
          <input type="date" value={filterDateFrom} onChange={e=>{setFilterDateFrom(e.target.value);if(e.target.value)setFilterSingleDate('');}} style={sty.dateInput} />
        </div>
        <div style={{flex:'1 1 110px',minWidth:'100px'}}>
          <label style={sty.label}><Calendar size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> To</label>
          <input type="date" value={filterDateTo} onChange={e=>{setFilterDateTo(e.target.value);if(e.target.value)setFilterSingleDate('');}} style={sty.dateInput} />
        </div>
      </div>

      {loading ? <div style={{padding:'50px',textAlign:'center',color:'var(--muted)'}}>Loading analytics...</div> : (
        <>
          {/* Summary Cards */}
          <div style={{display:'flex',flexWrap:'wrap',gap:'12px',marginBottom:'20px'}}>
            {[
              {icon:<TrendingUp size={16}/>,label:'Total Actions',value:rangedActivities.length,color:'#6366f1'},
              {icon:<User size={16}/>,label:'Active Users',value:userStats.filter(u=>u.totalActions>0).length,color:'#10b981'},
              {icon:<TrendingUp size={16}/>,label:'Top Worker',value:topWorker?.name||'-',color:'#f59e0b'},
              {icon:<LogIn size={16}/>,label:'Logins',value:rangedActivities.filter(a=>a.action==='login'||a.action==='admin_login').length,color:'#3b82f6'},
              {icon:<Edit3 size={16}/>,label:'Edits',value:rangedActivities.filter(a=>['edit_cells','add_row','delete_row','add_column','delete_column'].includes(a.action)).length,color:'#8b5cf6'},
              {icon:<Download size={16}/>,label:'Downloads',value:rangedActivities.filter(a=>a.action==='download_data').length,color:'#06b6d4'},
            ].map((c,i)=>(
              <div key={i} style={sty.statCard}>
                <div style={{display:'flex',alignItems:'center',gap:'8px',marginBottom:'6px'}}>
                  <div style={{width:'28px',height:'28px',borderRadius:'8px',background:`${c.color}15`,display:'flex',alignItems:'center',justifyContent:'center',color:c.color}}>{c.icon}</div>
                  <span style={{fontSize:'11px',color:'var(--muted)',fontWeight:700,textTransform:'uppercase',letterSpacing:'0.04em'}}>{c.label}</span>
                </div>
                <div style={{fontSize:'22px',fontWeight:800,color:'var(--foreground)'}}>{c.value}</div>
              </div>
            ))}
          </div>

          {/* Charts Row */}
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit, minmax(280px, 1fr))',gap:'16px',marginBottom:'20px'}}>
            {/* Daily Activity Bar Chart */}
            <div style={sty.card}>
              <BarChart data={dailyData} label="Daily Activity" />
            </div>
            {/* Action Distribution Donut */}
            <div style={sty.card}>
              <div style={{fontSize:'12px',fontWeight:700,color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.04em',marginBottom:'12px'}}>Action Distribution</div>
              <DonutChart data={actionDist} />
            </div>
          </div>

          {/* Second Charts Row */}
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit, minmax(280px, 1fr))',gap:'16px',marginBottom:'20px'}}>
            {/* Hourly Distribution */}
            <div style={sty.card}>
              <BarChart data={hourlyData} label="Activity by Hour (24h)" />
            </div>
            {/* Per-User Comparison */}
            <div style={sty.card}>
              <BarChart data={userBarData} label="Actions by Employee" />
            </div>
          </div>

          {/* Per-Employee Cards */}
          <h3 style={{fontSize:'13px',color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'12px',fontWeight:700}}>Per-Employee Breakdown</h3>
          <div style={{display:'flex',flexDirection:'column',gap:'12px'}}>
            {displayUsers.map((u,idx) => {
              const maxA = userStats[0]?.totalActions||1;
              const days = Object.keys(u.actionsByDay).length;
              const dayBars = Object.entries(u.actionsByDay).sort((a,b)=>a[0].localeCompare(b[0])).slice(-10)
                .map(([k,v])=>({key:new Date(k).toLocaleDateString('en-US',{month:'short',day:'numeric'}),value:v,color:`hsl(${(idx*50)%360},60%,55%)`}));
              return (
                <div key={u.id} style={sty.card}>
                  <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'12px',flexWrap:'wrap',gap:'8px'}}>
                    <div style={{display:'flex',alignItems:'center',gap:'12px'}}>
                      <div style={{width:'36px',height:'36px',borderRadius:'10px',background:`hsl(${(idx*50)%360},70%,95%)`,display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,fontSize:'14px',color:`hsl(${(idx*50)%360},60%,45%)`}}>
                        {u.name.charAt(0).toUpperCase()}
                      </div>
                      <div>
                        <div style={{fontSize:'15px',fontWeight:700,color:'var(--foreground)'}}>{u.name}</div>
                        <div style={{fontSize:'12px',color:'var(--muted)'}}>{u.email}</div>
                      </div>
                    </div>
                    <div style={{display:'flex',gap:'6px',flexWrap:'wrap'}}>
                      <span style={{padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:700,background:'rgba(99,102,241,0.1)',color:'#6366f1'}}>{u.totalActions} actions</span>
                      <span style={{padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:700,background:'rgba(16,185,129,0.1)',color:'#10b981'}}>{days} active day{days!==1?'s':''}</span>
                    </div>
                  </div>

                  {/* Stats grid */}
                  <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit, minmax(90px, 1fr))',gap:'8px',marginBottom:'12px'}}>
                    {[
                      {icon:<LogIn size={10}/>,label:'Logins',val:u.logins,color:'#3b82f6'},
                      {icon:<Edit3 size={10}/>,label:'Edits',val:u.edits,color:'#8b5cf6'},
                      {icon:<Download size={10}/>,label:'Downloads',val:u.downloads,color:'#f59e0b'},
                      {icon:<Shield size={10}/>,label:'Perms',val:u.permChanges,color:'#ef4444'},
                    ].map((s,i)=>(
                      <div key={i} style={{background:'var(--background)',borderRadius:'8px',padding:'10px 12px',border:'1px solid var(--border-light)'}}>
                        <div style={{fontSize:'11px',color:'var(--muted)',fontWeight:600,display:'flex',alignItems:'center',gap:'4px'}}>{s.icon} {s.label}</div>
                        <div style={{fontSize:'18px',fontWeight:800,color:s.color}}>{s.val}</div>
                      </div>
                    ))}
                  </div>

                  {/* Mini bar chart for this user's daily activity */}
                  {dayBars.length > 1 && (
                    <div style={{marginBottom:'10px'}}>
                      <BarChart data={dayBars} label={`${u.name}'s Daily Activity`} />
                    </div>
                  )}

                  {/* Volume bar */}
                  <div>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'4px'}}>
                      <span style={{fontSize:'11px',color:'var(--muted)',fontWeight:600}}>Work volume vs top</span>
                      <span style={{fontSize:'11px',color:'var(--navy)',fontWeight:700}}>{maxA>0?Math.round((u.totalActions/maxA)*100):0}%</span>
                    </div>
                    <div style={{width:'100%',height:'8px',background:'var(--border-light)',borderRadius:'4px',overflow:'hidden'}}>
                      <div style={{width:`${maxA>0?(u.totalActions/maxA)*100:0}%`,height:'100%',background:'var(--navy)',borderRadius:'4px',transition:'width 0.5s ease'}} />
                    </div>
                  </div>

                  {u.lastActive && (
                    <div style={{fontSize:'11px',color:'var(--muted)',marginTop:'8px',display:'flex',alignItems:'center',gap:'4px'}}>
                      <Clock size={10}/> Last active: {new Date(u.lastActive).toLocaleString()}
                    </div>
                  )}
                </div>
              );
            })}
            {displayUsers.length === 0 && <div style={{padding:'40px',textAlign:'center',color:'var(--muted)',background:'var(--surface)',borderRadius:'12px',border:'1px dashed var(--border)'}}>No employee data for the selected filters</div>}
          </div>
        </>
      )}
    </div>
  );
}
