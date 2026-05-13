import { useState, useEffect, useMemo } from 'react';
import { useAuth } from '../../lib/auth';
import { firebaseGetAllDownloadRequests, firebaseRespondDownloadRequest } from '../../lib/firebaseAuth';
import { Download, CheckCircle, XCircle, Clock, RefreshCw, Filter, User, ChevronDown, Send, MessageSquare, Calendar } from 'lucide-react';

export default function AdminDownloadRequestsPage() {
  const { user } = useAuth();
  const [requests, setRequests] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [respondingTo, setRespondingTo] = useState<string|null>(null);
  const [adminNote, setAdminNote] = useState('');
  const [filterStatus, setFilterStatus] = useState<string>('all');
  const [filterUser, setFilterUser] = useState<string>('all');

  const fetch_ = async () => {
    setLoading(true);
    try { const d = await firebaseGetAllDownloadRequests(); setRequests(d.requests||[]); }
    catch { } finally { setLoading(false); }
  };
  useEffect(() => { fetch_(); }, []);

  const handleRespond = async (id:string, status:'approved'|'rejected') => {
    if (status === 'rejected' && !adminNote.trim()) {
      alert('Please provide a reason for rejection');
      return;
    }
    try {
      await firebaseRespondDownloadRequest(id, status, adminNote, user?.name || user?.email || 'Admin');
      setRespondingTo(null); setAdminNote(''); fetch_();
    } catch(e:any) { alert(e.message); }
  };

  // Unique users from requests
  const requestUsers = useMemo(() => {
    const map = new Map<string, string>();
    requests.forEach(r => { if (r.userId && r.userName) map.set(r.userId, r.userName); });
    return Array.from(map.entries()).sort((a, b) => a[1].localeCompare(b[1]));
  }, [requests]);

  const filtered = useMemo(() => {
    return requests.filter(r => {
      if (filterStatus !== 'all' && r.status !== filterStatus) return false;
      if (filterUser !== 'all' && r.userId !== filterUser) return false;
      return true;
    });
  }, [requests, filterStatus, filterUser]);

  const pending = filtered.filter(r => r.status === 'pending');
  const past = filtered.filter(r => r.status !== 'pending');

  const sBadge = (s:string) => {
    if(s==='pending') return <span style={{padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:700,background:'rgba(245,158,11,0.1)',color:'#D97706',display:'flex',alignItems:'center',gap:'4px'}}><Clock size={12}/> Pending</span>;
    if(s==='approved') return <span style={{padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:700,background:'rgba(76,175,26,0.1)',color:'var(--brand-green)',display:'flex',alignItems:'center',gap:'4px'}}><CheckCircle size={12}/> Approved</span>;
    return <span style={{padding:'4px 10px',borderRadius:'6px',fontSize:'11px',fontWeight:700,background:'var(--destructive-bg)',color:'var(--destructive)',display:'flex',alignItems:'center',gap:'4px'}}><XCircle size={12}/> Rejected</span>;
  };

  return (
    <div>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:'16px',flexWrap:'wrap',gap:'10px'}}>
        <h2 style={{margin:0,fontSize:'18px',fontWeight:700,color:'var(--foreground)',display:'flex',alignItems:'center',gap:'10px'}}><Download size={20} color="var(--navy)"/> Download Requests</h2>
        <button onClick={fetch_} style={{background:'var(--surface)',border:'1px solid var(--border)',color:'var(--navy)',cursor:'pointer',padding:'10px',borderRadius:'8px',display:'flex',boxShadow:'var(--shadow-sm)'}}><RefreshCw size={16}/></button>
      </div>

      {/* Filters */}
      <div style={{
        display:'flex',flexWrap:'wrap',gap:'12px',marginBottom:'16px',
        background:'var(--surface)',borderRadius:'10px',padding:'12px 16px',
        border:'1px solid var(--border)'
      }}>
        <div style={{flex:'1 1 160px',minWidth:'140px'}}>
          <label style={{fontSize:'11px',fontWeight:700,color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'4px',display:'block'}}>
            <Filter size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> Status
          </label>
          <div style={{position:'relative'}}>
            <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
              style={{width:'100%',padding:'8px 28px 8px 10px',borderRadius:'6px',border:'1px solid var(--border)',background:'var(--background)',color:'var(--foreground)',fontSize:'13px',fontWeight:500,appearance:'none',cursor:'pointer',outline:'none'}}>
              <option value="all">All Status</option>
              <option value="pending">Pending</option>
              <option value="approved">Approved</option>
              <option value="rejected">Rejected</option>
            </select>
            <ChevronDown size={12} style={{position:'absolute',right:'8px',top:'50%',transform:'translateY(-50%)',pointerEvents:'none',color:'var(--muted)'}}/>
          </div>
        </div>
        <div style={{flex:'1 1 160px',minWidth:'140px'}}>
          <label style={{fontSize:'11px',fontWeight:700,color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'4px',display:'block'}}>
            <User size={10} style={{marginRight:'3px',verticalAlign:'middle'}}/> User
          </label>
          <div style={{position:'relative'}}>
            <select value={filterUser} onChange={e => setFilterUser(e.target.value)}
              style={{width:'100%',padding:'8px 28px 8px 10px',borderRadius:'6px',border:'1px solid var(--border)',background:'var(--background)',color:'var(--foreground)',fontSize:'13px',fontWeight:500,appearance:'none',cursor:'pointer',outline:'none'}}>
              <option value="all">All Users</option>
              {requestUsers.map(([id, name]) => <option key={id} value={id}>{name}</option>)}
            </select>
            <ChevronDown size={12} style={{position:'absolute',right:'8px',top:'50%',transform:'translateY(-50%)',pointerEvents:'none',color:'var(--muted)'}}/>
          </div>
        </div>
        <div style={{display:'flex',alignItems:'flex-end'}}>
          <span style={{fontSize:'12px',color:'var(--muted)',fontWeight:500,padding:'8px 0'}}>
            {filtered.length} result{filtered.length !== 1 ? 's' : ''}
          </span>
        </div>
      </div>

      {loading ? <div style={{padding:'50px',textAlign:'center',color:'var(--muted)'}}>Loading...</div> : (
        <div style={{display:'flex',flexDirection:'column',gap:'20px'}}>
          
          {/* Pending Requests */}
          {pending.length > 0 && (
            <div>
              <h3 style={{fontSize:'13px',color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'12px',fontWeight:700,display:'flex',alignItems:'center',gap:'6px'}}>
                <Clock size={14} color="#D97706"/> Pending Action ({pending.length})
              </h3>
              <div style={{display:'flex',flexDirection:'column',gap:'12px'}}>
                {pending.map(r => (
                  <div key={r.id} style={{background:'var(--surface)',borderRadius:'12px',border:'2px solid rgba(245,158,11,0.3)',padding:'16px 20px',boxShadow:'var(--shadow-md)'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:'12px',flexWrap:'wrap',gap:'8px'}}>
                      <div>
                        <div style={{fontSize:'15px',color:'var(--foreground)',fontWeight:700,marginBottom:'4px'}}>{r.registerName}</div>
                        <div style={{fontSize:'13px',color:'var(--muted)'}}>Requested by <strong style={{color:'var(--foreground)'}}>{r.userName}</strong></div>
                        <div style={{fontSize:'11px',color:'var(--muted)',marginTop:'2px',display:'flex',alignItems:'center',gap:'4px'}}>
                          <Calendar size={10}/> {new Date(r.createdAt).toLocaleString()}
                        </div>
                      </div>
                      {sBadge(r.status)}
                    </div>
                    <div style={{background:'var(--background)',padding:'12px',borderRadius:'8px',fontSize:'13px',color:'var(--foreground)',marginBottom:'16px',border:'1px solid var(--border)'}}>
                      <strong style={{color:'var(--navy)',fontSize:'11px',textTransform:'uppercase',letterSpacing:'0.03em'}}>Reason:</strong>
                      <div style={{marginTop:'4px'}}>{r.description}</div>
                    </div>
                    
                    {respondingTo === r.id ? (
                      <div style={{background:'var(--background)',padding:'16px',borderRadius:'10px',border:'1px solid var(--border)'}}>
                        <div style={{display:'flex',alignItems:'center',gap:'6px',marginBottom:'10px',fontSize:'13px',fontWeight:700,color:'var(--navy)'}}>
                          <MessageSquare size={14}/> Admin Response
                        </div>
                        <textarea
                          placeholder="Type your response or rejection reason here..."
                          value={adminNote}
                          onChange={e => setAdminNote(e.target.value)}
                          style={{
                            width:'100%',background:'var(--surface)',border:'1.5px solid var(--border)',
                            color:'var(--foreground)',fontSize:'14px',outline:'none',minHeight:'80px',
                            padding:'12px',borderRadius:'8px',resize:'vertical',boxSizing:'border-box',
                            fontFamily:'inherit'
                          }}
                        />
                        <div style={{display:'flex',gap:'10px',justifyContent:'flex-end',marginTop:'12px',flexWrap:'wrap'}}>
                          <button onClick={() => { setRespondingTo(null); setAdminNote(''); }} style={{padding:'10px 18px',borderRadius:'8px',border:'1px solid var(--border)',background:'var(--surface)',color:'var(--muted)',cursor:'pointer',fontWeight:600,fontSize:'13px'}}>Cancel</button>
                          <button onClick={() => handleRespond(r.id,'rejected')} style={{padding:'10px 18px',borderRadius:'8px',border:'none',background:'var(--destructive-bg)',color:'var(--destructive)',cursor:'pointer',fontWeight:700,fontSize:'13px',display:'flex',alignItems:'center',gap:'6px'}}>
                            <XCircle size={14}/> Reject
                          </button>
                          <button onClick={() => handleRespond(r.id,'approved')} style={{padding:'10px 18px',borderRadius:'8px',border:'none',background:'linear-gradient(135deg, var(--brand-green), var(--brand-green-dark))',color:'white',cursor:'pointer',fontWeight:700,fontSize:'13px',boxShadow:'var(--shadow-button)',display:'flex',alignItems:'center',gap:'6px'}}>
                            <CheckCircle size={14}/> Approve
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div style={{display:'flex',gap:'10px',flexWrap:'wrap'}}>
                        <button onClick={() => setRespondingTo(r.id)} style={{padding:'10px 20px',borderRadius:'8px',border:'1px solid var(--navy)',background:'var(--surface)',color:'var(--navy)',cursor:'pointer',fontSize:'13px',fontWeight:700,boxShadow:'var(--shadow-sm)',display:'flex',alignItems:'center',gap:'6px'}}>
                          <Send size={14}/> Respond
                        </button>
                        <button onClick={() => handleRespond(r.id,'approved')} style={{padding:'10px 20px',borderRadius:'8px',border:'none',background:'linear-gradient(135deg, var(--brand-green), var(--brand-green-dark))',color:'white',cursor:'pointer',fontSize:'13px',fontWeight:700,boxShadow:'var(--shadow-button)',display:'flex',alignItems:'center',gap:'6px'}}>
                          <CheckCircle size={14}/> Quick Approve
                        </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {pending.length === 0 && filterStatus !== 'pending' && (
            <div style={{padding:'20px',textAlign:'center',color:'var(--muted)',background:'var(--surface)',borderRadius:'12px',border:'1px dashed var(--border)',fontSize:'14px'}}>
              ✅ No pending requests
            </div>
          )}

          {/* Past Requests */}
          {past.length > 0 && (
            <div>
              <h3 style={{fontSize:'13px',color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:'12px',fontWeight:700}}>History ({past.length})</h3>
              <div style={{display:'flex',flexDirection:'column',gap:'8px'}}>
                {past.map(r => (
                  <div key={r.id} style={{background:'var(--surface)',borderRadius:'10px',border:'1px solid var(--border)',padding:'14px 18px',boxShadow:'var(--shadow-sm)'}}>
                    <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:'8px',flexWrap:'wrap'}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:'14px',color:'var(--foreground)',fontWeight:600}}>{r.registerName}</div>
                        <div style={{fontSize:'12px',color:'var(--muted)',marginTop:'2px',display:'flex',flexWrap:'wrap',gap:'4px',alignItems:'center'}}>
                          <span style={{fontWeight:600}}>{r.userName}</span>
                          <span>•</span>
                          <span>{new Date(r.createdAt).toLocaleDateString()}</span>
                        </div>
                      </div>
                      {sBadge(r.status)}
                    </div>
                    {r.description && (
                      <div style={{fontSize:'12px',color:'var(--foreground)',marginTop:'8px',paddingTop:'8px',borderTop:'1px solid var(--border-light)'}}>
                        <strong style={{color:'var(--muted)'}}>Request:</strong> {r.description}
                      </div>
                    )}
                    {r.adminResponse && (
                      <div style={{fontSize:'12px',color:'var(--foreground)',marginTop:'6px',background:'var(--background)',padding:'8px 10px',borderRadius:'6px',border:'1px solid var(--border-light)'}}>
                        <strong style={{color:'var(--navy)'}}>Admin Note:</strong> {r.adminResponse}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {filtered.length === 0 && <div style={{padding:'40px',textAlign:'center',color:'var(--muted)',fontSize:'14px'}}>No download requests found</div>}
        </div>
      )}
    </div>
  );
}
