<?php
use Cake\Event\EventManager;

EventManager::instance()
	->attach(
		function (Cake\Event\Event $event) {
			$controller = $event->subject();
			if ($controller->components()->loaded('RequestHandler')) {
				$controller->RequestHandler->viewClassMap('xlsx', 'Dakota/CakeExcel.Excel');
			}
		},
		'Controller.initialize'
	);
