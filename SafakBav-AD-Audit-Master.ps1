# =============================================================================
# PROJE: AD ANALIZ DASHBOARD v5.3 (FORTRESS)
# YAZAR: SAFAK CAN BAV
# =============================================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host '[+] AD Evrensel Analiz Baslatildi (v5.3 FORTRESS)...' -ForegroundColor Yellow

try {
    $root = [ADSI]"LDAP://RootDSE"
    $rootPath = "LDAP://" + $root.defaultNamingContext
} catch {
    $rootPath = "LDAP://" + $env:USERDNSDOMAIN
}

# --- SORGULAR ---
$uS = New-Object DirectoryServices.DirectorySearcher([ADSI]$rootPath)
$uS.Filter = '(&(objectCategory=person)(objectClass=user))'; $uS.PageSize = 25000; $uS.SizeLimit = 0
$uS.PropertiesToLoad.AddRange(@('samaccountname','displayname','department','useraccountcontrol','lastlogontimestamp','memberOf','distinguishedName','adminCount','title','description'))
$uRes = $uS.FindAll()

$cS = New-Object DirectoryServices.DirectorySearcher([ADSI]$rootPath)
$cS.Filter = '(objectClass=computer)'; $cS.PageSize = 25000; $cS.SizeLimit = 0
$cS.PropertiesToLoad.AddRange(@('name','operatingsystem','lastlogontimestamp','distinguishedName'))
$cRes = $cS.FindAll()

function Get-DeepPath($dn){
    if(!$dn -or $dn -eq ""){return "Genel"}
    try {
        $dnS = $dn.ToString()
        $parts = $dnS -split ","
        $res = ""
        foreach($p in $parts){
            if($p -like "OU=*"){
                $val = $p.Replace("OU=","")
                if($val -notmatch 'Users|Staff|Admin|Personel|Guest|Config|System|Disabled'){
                    if(!$res){ $res = $val } else { $res = "$val / $res" }
                }
            }
        }
        return if($res){$res}else{"Merkez"}
    } catch { return "Genel" }
}

$dom = "AD Domain"; if($env:USERDNSDOMAIN){ $dom = $env:USERDNSDOMAIN }
$uL = New-Object System.Collections.Generic.List[Object]
$cL = New-Object System.Collections.Generic.List[Object]
$s = @{ T1=0; T2=0; T3=0; T4=0; TC=0; DN=$dom }
$Today = Get-Date

Write-Host "[+] Veriler isleniyor... ($($uRes.Count) Kullanici)" -F Yellow
foreach($r in $uRes){
    $p = $r.Properties; $sam = if($p.samaccountname){$p.samaccountname[0].ToString()}else{""}
    if($sam -like 'HealthMailbox*' -or $sam -like 'SM_*' -or $sam -eq 'Guest'){continue}
    $s.T1++; $uac = if($p.useraccountcontrol){[int]$p.useraccountcontrol[0]}else{512}
    $en = ($uac -band 2) -eq 0; $adm = if(($p.adminCount -and $p.adminCount[0] -eq 1) -or ($p.memberOf -match 'Admin')){ $true } else { $false }
    if($adm){$s.T4++}
    
    $uT = "-"; if($p.title){$uT=$p.title[0].ToString()}elseif($p.description){$uT=$p.description[0].ToString()}
    
    # SAFE DN ACCESS
    $dnVal = ""
    if($p.distinguishedName -and $p.distinguishedName.Count -gt 0){ $dnVal = $p.distinguishedName[0].ToString() }
    $uD = if($p.department){$p.department[0].ToString()}else{Get-DeepPath($dnVal)}

    $ll = 'Never'; $st = 'Never'
    if($p.lastlogontimestamp){ try{ $dt=[datetime]::FromFileTime($p.lastlogontimestamp[0]); $ll=$dt.ToString('dd.MM.yy'); $df=($Today-$dt).Days; if($df -le 90){$st='Aktif'; if($en){$s.T2++}}else{$st='Atil'; if($en){$s.T3++}} }catch{} }
    $uL.Add(@{ A=if($p.displayname){$p.displayname[0].ToString()}else{$sam}; K=$sam; D=$uD; J=$uT; S=if($en){'Aktif'}else{'Pasif'}; U=$st; L=$ll; M=if($adm){1}else{0} })
}

foreach($r in $cRes){
    $p = $r.Properties; $s.TC++; $ll = 'Never'; $st = 'Never'
    if($p.lastlogontimestamp){ try{ $dt=[datetime]::FromFileTime($p.lastlogontimestamp[0]); $ll=$dt.ToString('dd.MM.yy'); $df=($Today-$dt).Days; if($df -le 90){$st='Aktif'}else{$st='Atil'} }catch{} }
    $dnComp = ""; if($p.distinguishedName -and $p.distinguishedName.Count -gt 0){ $dnComp = $p.distinguishedName[0].ToString() }
    $cL.Add(@{ N=if($p.name){$p.name[0].ToString()}else{"-"}; O=if($p.operatingsystem){$p.operatingsystem[0].ToString()}else{"-"}; K=Get-DeepPath($dnComp); S=$st; L=$ll } )
}

$uEnc = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($uL | ConvertTo-Json -Depth 2 -Compress)))
$cEnc = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($cL | ConvertTo-Json -Depth 2 -Compress)))
$sEnc = ($s | ConvertTo-Json -Compress)

$html = @'
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>AD AUDIT MASTER</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
    <style>
        body { background:#f8fafc; font-family:'Plus Jakarta Sans',sans-serif; font-size:11px; }
        .sidebar { width:180px; height:100vh; position:fixed; background:#fff; border-right:1px solid #e2e8f0; padding:20px; }
        .main { margin-left:180px; padding:25px; }
        .card-p { background:#fff; border-radius:12px; border:1px solid #e2e8f0; padding:15px; margin-bottom:20px; box-shadow:0 1px 2px rgba(0,0,0,0.05); }
        .stat-val { font-size:22px; font-weight:700; color:#1e293b; }
        .badge-s { padding:4px 10px; border-radius:50px; font-weight:700; font-size:0.6rem; }
        .nav-link { border:none !important; color:#64748b; font-weight:600; border-radius:8px; margin-bottom:5px; padding:8px 12px; cursor:pointer; }
        .nav-link.active { background:#eff6ff !important; color:#2563eb !important; }
    </style>
</head>
<body>
    <div class="sidebar d-flex flex-column">
        <h5 class="fw-bold mb-5 text-primary">AD ANAL&#304;Z</h5>
        <div class="nav flex-column nav-pills">
            <div class="nav-link active" onclick="tab('t1', this)">Dashboard</div>
            <div class="nav-link" onclick="tab('t2', this)">Kullan&#305;c&#305;lar</div>
            <div class="nav-link" onclick="tab('t3', this)">Bilgisayarlar</div>
        </div>
    </div>
    <div class="main">
        <h2 class="fw-bold mb-4">AD AUDIT DASHBOARD</h2>
        <div id="t1" class="tab-pane active shadow-none">
            <div class="row g-3 text-center">
                <div class="col-md-3"><div class="card-p">USERS<div class="stat-val" id="v1">0</div></div></div>
                <div class="col-md-3"><div class="card-p text-success">ACTIVE<div class="stat-val" id="v2">0</div></div></div>
                <div class="col-md-3"><div class="card-p text-danger">STALE<div class="stat-val" id="v3">0</div></div></div>
                <div class="col-md-3"><div class="card-p text-primary">ADMINS<div class="stat-val" id="v4">0</div></div></div>
            </div>
            <div class="card-p mt-4"><canvas id="c1" height="150"></canvas></div>
        </div>
        <div id="t2" class="tab-pane d-none"><div class="card-p"><table class="table w-100" id="ut"><thead><tr><th>&#304;sim</th><th>Kullan&#305;c&#305;</th><th>&#220;nvan</th><th>Kategori</th><th>Status</th><th>Giri&#351;</th></tr></thead><tbody id="ub"></tbody></table></div></div>
        <div id="t3" class="tab-pane d-none"><div class="card-p"><table class="table w-100" id="ct"><thead><tr><th>PC</th><th>OS</th><th>Kategori</th><th>Status</th><th>Giri&#351;</th></tr></thead><tbody id="cb"></tbody></table></div></div>
    </div>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        const uB = "@U@"; const cB = "@C@"; const s = @S@;
        function fix64(b){ const s = atob(b); const u8 = new Uint8Array(s.length); for(let i=0; i<s.length; i++) u8[i] = s.charCodeAt(i); return JSON.parse(new TextDecoder().decode(u8)); }
        window.onload = function(){
            const u = fix64(uB); const c = fix64(cB);
            $('#v1').text(s.T1); $('#v2').text(s.T2); $('#v3').text(s.T3); $('#v4').text(s.T4);
            $('#ub').html(u.map(x=>`<tr><td><b>${x.M?'<span class="text-danger">👑 </span>':''}${x.A}</b></td><td><code>${x.K}</code></td><td>${x.J}</td><td>${x.D}</td><td><span class="badge-s ${x.S=='Aktif'?'bg-success text-white':'bg-danger text-white'}">${x.S}</span></td><td>${x.L}</td></tr>`).join(''));
            $('#cb').html(c.map(x=>`<tr><td><b>${x.N}</b></td><td>${x.O}</td><td>${x.K}</td><td><span class="badge-s ${x.S=='Aktif'?'bg-success text-white':'bg-danger text-white'}">${x.S}</span></td><td>${x.L}</td></tr>`).join(''));
            $('#ut,#ct').DataTable({pageLength:25});
            const d = u.reduce((a,b)=>{a[b.D]=(a[b.D]||0)+1;return a;},{});
            new Chart($('#c1'),{type:'bar',data:{labels:Object.keys(d),datasets:[{data:Object.values(d),backgroundColor:'#3b82f6',borderRadius:10}]}});
        };
        function tab(id, el){ $('.tab-pane').addClass('d-none'); $('#'+id).removeClass('d-none'); $('.nav-link').removeClass('active'); $(el).addClass('active'); }
    </script>
</body>
</html>
'@

$f = $html.Replace('@U@',$uEnc).Replace('@C@',$cEnc).Replace('@S@',$sEnc)
[IO.File]::WriteAllText([IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'AD_FORTRESS_REPORT.html'), $f, [Text.Encoding]::UTF8)
Write-Host '--- FİNAL: AD_FORTRESS_REPORT.html ---' -ForegroundColor Green
