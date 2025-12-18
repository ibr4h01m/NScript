var nsapi = (function () {
  function create(progId) {
    return new ActiveXObject(String(progId));
  }

  function call(progId, methodName) {
    var obj = create(progId);
    var args = [];
    for (var i = 2; i < arguments.length; i++) args.push(arguments[i]);
    return obj[String(methodName)].apply(obj, args);
  }

  return {
    create: create,
    call: call
  };
})();
