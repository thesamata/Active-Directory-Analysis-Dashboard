# =============================================================================
# PROJE: AD ANALIZ PANELI
# YAZAR: SAFAK CAN BAV
# =============================================================================
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host '[+] AD Analiz Paneli Baslatiliyor...' -ForegroundColor Yellow

try {
    $root = [ADSI]"LDAP://RootDSE"
    $rootPath = "LDAP://" + $root.defaultNamingContext
} catch { $rootPath = "LDAP://" + $env:USERDNSDOMAIN }

# --- VERI TOPLAMA (DEGISMEDEN KORUNDU) ---
$uS = New-Object DirectoryServices.DirectorySearcher([ADSI]$rootPath)
$uS.Filter = '(&(objectCategory=person)(objectClass=user))'; $uS.PageSize = 2000; $uS.SizeLimit = 0
$uS.PropertiesToLoad.AddRange(@('samaccountname','displayname','department','useraccountcontrol','lastlogontimestamp','memberOf','distinguishedName','adminCount','title','description','pwdlastset','msDS-UserPasswordExpiryTimeComputed','manager','lockoutTime','whenCreated','mail','primaryGroupID'))
$uRes = $uS.FindAll()

$cS = New-Object DirectoryServices.DirectorySearcher([ADSI]$rootPath)
$cS.Filter = '(objectClass=computer)'; $cS.PageSize = 2000; $cS.SizeLimit = 0
$cS.PropertiesToLoad.AddRange(@('name','operatingsystem','lastlogontimestamp','distinguishedName'))
$cRes = $cS.FindAll()

function Get-DeepPath($dn){
    if(!$dn -or $dn -eq ""){return "Genel"}
    try {
        $dnS = $dn.ToString()
        $parts = $dnS -split ","
        $resList = New-Object System.Collections.Generic.List[string]
        foreach($p in $parts){
            if($p -like "OU=*"){
                $val = $p.Replace("OU=","")
                if($val -notmatch 'Users|Staff|Admin|Personel|Guest|Config|System|Disabled'){ $resList.Add($val) }
            }
        }
        return if($resList.Count -gt 0){ [string]::Join(" / ", $resList) }else{"Merkez"}
    } catch { return "Genel" }
}

$dom = "AD Domain"; if($env:USERDNSDOMAIN){ $dom = $env:USERDNSDOMAIN }
$uL = New-Object System.Collections.Generic.List[Object]
$cL = New-Object System.Collections.Generic.List[Object]
$Today = Get-Date
$s = @{ T1=0; T2=0; T3=0; T4=0; T5=0; T6=0; TC=0; DN=$dom }

Write-Host "[+] Veriler Isleniyor... ($($uRes.Count) Kullanici)" -F Cyan
foreach($r in $uRes){
    $p = $r.Properties; $sam = if($p.samaccountname){$p.samaccountname[0].ToString()}else{""}
    if($sam -like 'HealthMailbox*' -or $sam -like 'SM_*' -or $sam -eq 'Guest'){continue}
    $s.T1++;
    $uac = if($p.useraccountcontrol){[int]$p.useraccountcontrol[0]}else{512}
    $en = ($uac -band 2) -eq 0; 
    $lck = if($p.lockoutTime -and $p.lockoutTime[0] -gt 0){$true}else{$false}
    if($lck){$s.T6++}
    $adm = if(($p.adminCount -and $p.adminCount[0] -eq 1) -or ($p.memberOf -match 'Admin|Domain Admins|Enterprise Admins|Schema Admins|Account Operators|Backup Operators|Server Operators') -or ($p.primaryGroupID -and $p.primaryGroupID[0] -eq 512)){ $true } else { $false }
    if($adm){$s.T4++}
    $uT = "-"; if($p.title){$uT=$p.title[0].ToString()}elseif($p.description){$uT=$p.description[0].ToString()}
    $mgr = if($p.manager){ ($p.manager[0].ToString() -split ",")[0].Replace("CN=","") } else { "-" }
    $desc = if($p.description){$p.description[0].ToString()}else{""}
    $cre = if($p.whenCreated){ ([datetime]$p.whenCreated[0]).ToString('dd.MM.yyyy') } else {"-"}
    $mail = if($p.mail){$p.mail[0].ToString()}else{"-"}
    $dnVal = ""; if($p.distinguishedName){ $dnVal = $p.distinguishedName[0].ToString() }
    $uD = if($p.department){$p.department[0].ToString()}else{Get-DeepPath($dnVal)}
    $ll = 'Hic'; $st = 'Asla'; $diff = 999
    if($p.lastlogontimestamp){ 
        try{ 
            $dt=[datetime]::FromFileTime($p.lastlogontimestamp[0])
            if($dt.Year -lt 2100 -and $dt.Year -gt 1900){
                $ll=$dt.ToString('dd.MM.yy'); $diff=($Today-$dt).Days
                if($diff -le 90){$st='Aktif'; if($en){$s.T2++}}else{$st='Atil'; if($en){$s.T3++}} 
            }
        }catch{} 
    }
    $pwdStat = "Gecerli"; if($p.'msds-userpasswordexpirytimecomputed' -and $p.'msds-userpasswordexpirytimecomputed'[0] -gt 0){
        try{ if([datetime]::FromFileTime($p.'msds-userpasswordexpirytimecomputed'[0]) -lt $Today){$pwdStat="Biten"; $s.T5++} }catch{}
    }
    $grps = if($p.memberOf){ @($p.memberOf | ForEach-Object { ($_ -split ",")[0].Replace("CN=","") }) -join "|" } else { "" }
    $uL.Add(@{ A=if($p.displayname){$p.displayname[0].ToString()}else{$sam}; K=$sam; D=$uD; J=$uT; S=if($lck){'Kilitli'}elseif($en){'Aktif'}else{'Pasif'}; U=$st; L=$ll; M=if($adm){1}else{0}; P=$pwdStat; R=$mgr; X=$desc; C=$cre; G=$diff; GR=$grps; E=$mail })
}
foreach($r in $cRes){
    $p = $r.Properties; $s.TC++; $ll = 'Hic'; $st = 'Asla'
    if($p.lastlogontimestamp){ try{ $dt=[datetime]::FromFileTime($p.lastlogontimestamp[0]); $ll=$dt.ToString('dd.MM.yy'); $df=($Today-$dt).Days; if($df -le 90){$st='Aktif'}else{$st='Atil'} }catch{} }
    $dnComp = ""; if($p.distinguishedName){ $dnComp = $p.distinguishedName[0].ToString() }
    $cL.Add(@{ N=if($p.name){$p.name[0].ToString()}else{"-"}; O=if($p.operatingsystem){$p.operatingsystem[0].ToString()}else{"-"}; K=Get-DeepPath($dnComp); S=$st; L=$ll } )
}
$deptGroup = $uL | Group-Object D | Sort-Object Count -Descending
$s.DL = @($deptGroup | ForEach-Object {$_.Name}); $s.DC = @($deptGroup | ForEach-Object {$_.Count})
$uEnc = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($uL | ConvertTo-Json -Depth 2 -Compress)))
$cEnc = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($cL | ConvertTo-Json -Depth 2 -Compress)))
$sEnc = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($s | ConvertTo-Json -Compress)))

$html = @'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>AD Analiz Paneli</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root { --bg: #030712; --side: #0e121b; --card: #151921; --primary: #3b82f6; --text: #f1f5f9; --border: #232a35; }
        body { background: var(--bg); font-family: 'Plus Jakarta Sans', sans-serif; color: var(--text); overflow: hidden; display: flex; height: 100vh; margin:0; }
        .sidebar { width: 260px; background: var(--side); border-right: 1px solid var(--border); display: flex; flex-direction: column; padding: 25px 0; }
        .logo-box { padding: 0 25px 30px; border-bottom: 1px solid var(--border); margin-bottom: 25px; }
        .logo-box b { font-size: 1.1rem; color: #fff; letter-spacing: 1px; display: flex; align-items: center; gap: 10px; }
        .nav-link { padding: 12px 25px; color: #94a3b8; cursor: pointer; transition: 0.2s; display: flex; align-items: center; gap: 12px; font-weight: 600; border-right: 3px solid transparent; }
        .nav-link:hover { background: rgba(255,255,255,0.02); color: #fff; }
        .nav-link.active { color: #fff; background: rgba(59, 130, 246, 0.05); border-right-color: var(--primary); }
        .main-content { flex: 1; overflow-y: auto; padding: 30px; position: relative; }
        .top-info { position: absolute; top: 30px; right: 30px; color: #64748b; font-size: 11px; font-weight: 700; z-index: 5; }
        .stat-banner { display: grid; grid-template-columns: repeat(6, 1fr); gap: 20px; margin-bottom: 35px; }
        .stat-card { background: var(--card); padding: 20px; border-radius: 18px; border: 1px solid var(--border); cursor: pointer; transition: 0.3s; }
        .stat-card:hover { transform: translateY(-5px); border-color: var(--primary); }
        .stat-val { font-size: 26px; font-weight: 800; color: #fff; line-height: 1; }
        .stat-lab { font-size: 10px; font-weight: 700; color: #64748b; margin-top: 8px; text-transform: uppercase; }
        .card-box { background: var(--card); border-radius: 20px; border: 1px solid var(--border); padding: 25px; }
        .card-ttl { font-size: 15px; font-weight: 700; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; color: #fff; }
        .table { width: 100% !important; background: transparent !important; color: #cbd5e1 !important; border-collapse: collapse !important; border-color: var(--border) !important; }
        .table thead th { background: rgba(255,255,255,0.03) !important; color: #64748b !important; font-size: 10px !important; text-transform: uppercase !important; border-bottom: 2px solid var(--border) !important; padding: 15px !important; }
        .table tbody tr { background: transparent !important; }
        .table tbody td { background: transparent !important; padding: 12px 15px !important; border-bottom: 1px solid var(--border) !important; color: #cbd5e1 !important; vertical-align: middle !important; }
        .table.dataTable.no-footer { border-bottom: 1px solid var(--border) !important; }
        
        .badge-x { padding: 5px 10px; border-radius: 8px; font-weight: 700; font-size: 10px; display: inline-block; white-space: nowrap; }
        .bg-aktif { background: rgba(34, 197, 94, 0.1); color: #4ade80; }
        .bg-atil { background: rgba(249, 115, 22, 0.1); color: #fb923c; }
        .bg-pasif { background: rgba(148, 163, 184, 0.1); color: #94a3b8; }
        .bg-kilitli { background: rgba(239, 68, 68, 0.1); color: #f87171; }
        .bg-biten { background: rgba(239, 68, 68, 0.1); color: #f87171; }
        .bg-gecerli { background: rgba(34, 197, 94, 0.1); color: #4ade80; }
        .filter-inp { background: #0b0e14; border: 1px solid var(--border); color: #fff; border-radius: 8px; padding: 6px 12px; font-size: 11px; width: 100%; border-style: solid; border-width: 1px; }
        .modal-content { background: var(--side); border-radius: 24px; border: 1px solid var(--border); color: #fff; }
        .dataTables_filter input { background: #0b0e14; border: 1px solid var(--border); border-radius: 8px; color: #fff; padding: 6px 12px; margin-left: 10px; }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="logo-box"><b><i class="fa-solid fa-cube text-primary"></i> ANALIZ PANELI</b></div>
        <div class="nav-link active" onclick="goto('t1', this)"><i class="fa-solid fa-house"></i> Ozet Paneli</div>
        <div class="nav-link" onclick="goto('t2', this)"><i class="fa-solid fa-user-group"></i> Kullanicilar</div>
        <div class="nav-link" onclick="goto('t3', this)"><i class="fa-solid fa-laptop"></i> Cihazlar</div>
        <div class="mt-auto p-4"><small style="color:#475569; font-size:10px;">v2.0</small></div>
    </div>

    <main class="main-content">
        <div class="top-info" id="clock">00.00.0000 00:00</div>
        <div id="t1" class="tab-pane">
            <div class="stat-banner">
                <div class="stat-card" onclick="fStat('all')"><div class="stat-val" id="s1">0</div><div class="stat-lab">Toplam Kayit</div></div>
                <div class="stat-card" onclick="fStat('Aktif', 5)"><div class="stat-val text-success" id="s2">0</div><div class="stat-lab">Aktifler</div></div>
                <div class="stat-card" onclick="fStat('Atil', 6)"><div class="stat-val text-warning" id="s3">0</div><div class="stat-lab">Atillar</div></div>
                <div class="stat-card" onclick="fStat('Biten', 7)"><div class="stat-val text-danger" id="s5">0</div><div class="stat-lab">Parolasi Biten</div></div>
                <div class="stat-card" onclick="fStat('1', 11)"><div class="stat-val" style="color:#a78bfa" id="s4">0</div><div class="stat-lab">Sistem Admin</div></div>
                <div class="stat-card" onclick="fStat('Kilitli', 5)"><div class="stat-val" style="color:#fb7185" id="s6">0</div><div class="stat-lab">Kilitliler</div></div>
            </div>
            <div class="row g-4">
                <div class="col-md-7"><div class="card-box"><div class="card-ttl">Departman Verileri</div><div style="height:400px; overflow-y:auto;"><canvas id="cBar"></canvas></div></div></div>
                <div class="col-md-5"><div class="card-box"><div class="card-ttl">Hesap Durumu</div><div style="height:400px; position:relative;"><canvas id="cPie"></canvas><div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); text-align:center;"><div style="font-size:36px; font-weight:800;" id="pVal">0</div><div style="font-size:10px; color:#64748b;">KULLANICI</div></div></div></div></div>
            </div>
        </div>
        <div id="t2" class="tab-pane d-none">
            <div class="card-box">
                <div class="d-flex justify-content-between mb-4"><h5>Kullanici Database</h5><button class="btn btn-sm btn-outline-secondary" onclick="clearF()">Temizle</button></div>
                <table id="uT" class="table w-100">
                    <thead>
                        <tr><th>Isim</th><th>Kullanici</th><th>Unvan</th><th>Departman</th><th>Yonetici</th><th>Durum</th><th>Giris</th><th>Parola</th><th>Tarih</th><th>Gun</th><th>M</th><th>Adm</th></tr>
                        <tr class="f-row">
                            <th><input class="filter-inp" placeholder="ara.."></th>
                            <th><input class="filter-inp" placeholder="ara.."></th>
                            <th><input class="filter-inp" placeholder="ara.."></th>
                            <th><input class="filter-inp" placeholder="ara.."></th>
                            <th><input class="filter-inp" placeholder="ara.."></th>
                            <th><select class="filter-inp"><option value="">Hepsi</option><option value="Aktif">Aktif</option><option value="Kilitli">Kilitli</option></select></th>
                            <th><select class="filter-inp"><option value="">Hepsi</option><option value="Aktif">Aktif</option><option value="Atil">Atil</option></select></th>
                            <th><select class="filter-inp"><option value="">Hepsi</option><option value="Gecerli">Gecerli</option><option value="Biten">Biten</option></select></th>
                            <th></th><th><input class="filter-inp" placeholder=">90"></th><th></th><th></th>
                        </tr>
                    </thead>
                    <tbody id="uB"></tbody>
                </table>
            </div>
        </div>
        <div id="t3" class="tab-pane d-none">
            <div class="card-box"><h5 class="mb-4">Cihaz Envanteri</h5><table id="cT" class="table w-100"><thead><tr><th>Hostname</th><th>Isletim Sistemi</th><th>Konum</th><th>Durum</th><th>Giris</th></tr></thead><tbody id="cB"></tbody></table></div>
        </div>
    </main>
    <div class="modal fade" id="uM" tabindex="-1"><div class="modal-dialog modal-dialog-centered modal-lg"><div class="modal-content"><div class="modal-header border-0"><h5 class="fw-bold" id="mT"></h5><button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button></div><div class="modal-body"><div class="row g-4"><div class="col-md-6"><small class="text-muted fw-bold">E-POSTA</small><div id="mE" class="fw-bold text-primary"></div></div><div class="col-md-6"><small class="text-muted fw-bold">YONETICI</small><div id="mR"></div></div><div class="col-md-12"><small class="text-muted fw-bold mb-2 d-block">AD GRUPLARI</small><div id="mG" class="d-flex flex-wrap gap-1"></div></div></div></div></div></div></div>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        const uE = "@U@"; const cE = "@C@"; const sE = "@S@";
        let u, c, s, ut, ct;
        function dec(b){ try{ let r = atob(b); let u8 = new Uint8Array(r.length); for(let i=0; i<r.length; i++) u8[i] = r.charCodeAt(i); return JSON.parse(new TextDecoder().decode(u8)); }catch(e){ return []; } }
        window.onload = function(){
            u = dec(uE); c = dec(cE); s = dec(sE);
            $('#clock').text(new Date().toLocaleString('tr-TR'));
            if(s){
                $('#s1').text(s.T1); $('#s2').text(s.T2); $('#s3').text(s.T3); $('#s4').text(s.T4); $('#s5').text(s.T5); $('#s6').text(s.T6); $('#pVal').text(s.T1);
                new Chart($('#cBar'), { type: 'bar', data: { labels: s.DL, datasets: [{ data: s.DC, backgroundColor: '#3b82f6', borderRadius: 8 }] }, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins:{legend:{display:false}}, scales:{x:{display:false},y:{ticks:{color:'#64748b', font:{weight:'600'}}}} } });
                new Chart($('#cPie'), { type: 'doughnut', data: { labels: ['Aktif','Atil','Asla'], datasets: [{ data: [s.T2, s.T3, s.T1-s.T2-s.T3], backgroundColor: ['#22c55e', '#f97316', '#1a202c'], borderWidth: 0 }] }, options: { cutout:'78%', plugins:{legend:{display:false}} } });
            }
            $('#uB').html(u.map((x,i) => `<tr><td><a href="#" class="text-primary text-decoration-none fw-bold" onclick="viewU(${i})">${x.M?'<i class="fa-solid fa-crown text-warning"></i> ':''}${x.A}</a></td><td><code>${x.K}</code></td><td>${x.J}</td><td>${x.D}</td><td>${x.R}</td><td><span class="badge-x bg-${x.S.toLowerCase()}">${x.S}</span></td><td><span class="badge-x bg-${x.U.toLowerCase()}">${x.U}</span></td><td><span class="badge-x bg-${x.P.toLowerCase()}">${x.P}</span></td><td>${x.L}</td><td>${x.G}</td><td>${x.M}</td><td>${x.M}</td></tr>`).join(''));
            $('#cB').html(c.map(x => `<tr><td><b>${x.N}</b></td><td>${x.O}</td><td>${x.K}</td><td><span class="badge-x bg-${x.S.toLowerCase()}">${x.S}</span></td><td>${x.L}</td></tr>`).join(''));
            let l = { sSearch: "Hizli Ara:", sLengthMenu: "_MENU_", oPaginate: { sNext: ">>", sPrevious: "<<" } };
            ut = $('#uT').DataTable({ language: l, pageLength: 25, dom: 'frtip', columnDefs:[{targets:[10,11], visible:false}] });
            ct = $('#cT').DataTable({ language: l, pageLength: 25, dom: 'frtip' });
            $.fn.dataTable.ext.search.push(function(set, dat, idx) {
                if (set.nTable.id !== 'uT') return true;
                let active = true;
                $('.f-row input').each(function() {
                    let v = $(this).val().trim(); if (!v) return;
                    let ci = $(this).parent().index(); let cv = dat[ci] || "";
                    let m = v.match(/^([><]=?|==)\s*(\d+)$/);
                    if (m) {
                        let op = m[1]; let target = parseInt(m[2]); let cur = parseInt(cv.replace(/[^\d]/g,'')) || 0;
                        if (op === '>') active = active && (cur > target); else if (op === '<') active = active && (cur < target);
                    } else { if (cv.toLowerCase().indexOf(v.toLowerCase()) === -1) active = false; }
                });
                return active;
            });
            $('.filter-inp').on('click', e => e.stopPropagation());
            $('.f-row input, .f-row select').on('keyup change', () => { ut.draw(); });
            $('.cf-row input').on('keyup change', function() { ct.column($(this).parent().index()).search(this.value).draw(); });
        };
        function goto(id, el){ $('.nav-link').removeClass('active'); $(el).addClass('active'); $('.tab-pane').addClass('d-none'); $('#'+id).removeClass('d-none'); }
        function viewU(i){ let x = u[i]; $('#mT').text(x.A); $('#mE').text(x.E); $('#mR').text(x.R); $('#mG').html(x.GR ? x.GR.split('|').map(n => `<span class="badge-x" style="background:rgba(255,255,255,0.02); border:1px solid #232a35; margin:3px;">${n}</span>`).join('') : '-'); new bootstrap.Modal($('#uM')).show(); }
        function fStat(v, ci){ goto('t2', $('.nav-link')[1]); $('.filter-inp').val(''); if(v!=='all') ut.column(ci).search(v).draw(); else ut.search('').columns().search('').draw(); }
        function clearF(){ $('.filter-inp').val(''); ut.search('').columns().search('').draw(); ct.search('').columns().search('').draw(); }
    </script>
</body>
</html>
'@

$dtS = Get-Date -Format "dd.MM.yyyy"
$fP = [IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), "AD-Analiz-Raporu_$dtS.html")
$finalH = $html.Replace('@U@',$uEnc).Replace('@C@',$cEnc).Replace('@S@',$sEnc)
[IO.File]::WriteAllText($fP, $finalH, [System.Text.Encoding]::UTF8)
Write-Host "`n--- Rapor Olusturuldu: $fP ---" -ForegroundColor Green
