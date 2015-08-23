

rr  = $ws.range('B2',"F8")
rr.value = Time.now.to_s


sleep 3

$log.info $ws.name
$log.error "???'"
$log.debug 'debing'
$log.info 333
