<?php
class ControllerSettingPrices extends Controller {
    private $error = array();

    public function index() {
        $this->document->setTitle('Загрузка прайс-листов');

        if (($this->request->server['REQUEST_METHOD'] == 'POST')) {

            $price_conf = [
                'wr' => [
                    'files' => ['prices', 'remains'],
                    'name'  => 'Weltwasser / River',
                ],
                'ab' => [
                    'files' => ['price_ob','price_sm','remains'],
                    'name'  => 'Allen Brau',
                ],
                'rgw' => [
                    'files' => ['prices', 'remains'],
                    'name'  => 'RGW',
                ],
                'santi' => [
                    'files' => ['prices', 'remains'],
                    'name'  => 'SantiLine',
                ],
                'aquanet' => [
                    'files' => ['prices_m', 'remains_m', 'prices_v', 'remains_v'],
                    'name'  => 'Aquanet',
                ],
            ];

            $warning =  $success = [];

            foreach ($price_conf as $prefix => $conf) {
                $missing = [];
                $upload_dir = DIR_PPROC. "prices/{$prefix}/";
                $prefix .= '_';

                foreach ($conf['files'] as $file_key) {
                    if (!empty($_FILES[$prefix.$file_key]['name']) && is_uploaded_file($_FILES[$prefix.$file_key]['tmp_name'])) {

                        $filename = $file_key.'.xls';
                        $target_path = $upload_dir . $filename;

                        if (!move_uploaded_file($_FILES[$prefix.$file_key]['tmp_name'], $target_path)) {
                            $this->error[$prefix.$file_key] = 'Ошибка при сохранении файла ' . $filename;
                        }
                    } else {
                        $this->error[$prefix.$file_key] = 'Необходимо указать файл';
                        $missing[] = $file_key;
                    }
                }
                if (!$missing) {
                    //помечаем поставщика, как загруженного корректно
                    file_put_contents($upload_dir.'ready_flag', date());
                    $success[] = $conf['name'];
                } else {
                    $warning[] = $conf['name'];
                }
            }

            if (!$warning) {
                $this->session->data['success'] = 'Все файлы успешно загружены';
                $this->response->redirect($this->url->link('setting/prices', 'user_token=' . $this->session->data['user_token'], true));
            } else {
                $this->error['warning'] = 'Файлы некоторых поставщиков ('.implode(', ', $warning).') не были загружены!';
                if ($success) {
                    $this->session->data['success'] = 'Данные некоторых поставщиков ('.implode(', ', $success).') отправлены в обработку.';;
                }
            }
        }

        $fields = ['warning'];
        foreach ($price_conf as $prefix => $conf) {
            foreach ($conf['files'] as $file) {$fields[] = $prefix.'_'.$file;}
        }
        foreach ($fields as $field) {
            $data['error_'.$field] = $this->error[$field] ?? false;
        }
        $data['success'] = $this->session->data['success'] ?? false;
        $this->session->data['success'] = false;

        $data['cancel'] = $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'], true);

        $data['header'] = $this->load->controller('common/header');
        $data['column_left'] = $this->load->controller('common/column_left');
        $data['footer'] = $this->load->controller('common/footer');

        $this->response->setOutput($this->load->view('setting/prices', $data));
    }

}
